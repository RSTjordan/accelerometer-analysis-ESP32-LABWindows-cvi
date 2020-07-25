#include "excel2000.h"

static void CVIFUNC RefreshEventsRegOnBeforeRefresh_EventVTableFunc (void *thisPtr,
                                                                     VBOOL *cancel);

static void CVIFUNC RefreshEventsRegOnAfterRefresh_EventVTableFunc (void *thisPtr,
                                                                    VBOOL success);

static void CVIFUNC AppEventsRegOnNewWorkbook_EventVTableFunc (void *thisPtr,
                                                               LPDISPATCH wb);

static void CVIFUNC AppEventsRegOnSheetSelectionChange_EventVTableFunc (void *thisPtr,
                                                                        LPDISPATCH sh,
                                                                        LPDISPATCH target);

static void CVIFUNC AppEventsRegOnSheetBeforeDoubleClick_EventVTableFunc (void *thisPtr,
                                                                          LPDISPATCH sh,
                                                                          LPDISPATCH target,
                                                                          VBOOL *cancel);

static void CVIFUNC AppEventsRegOnSheetBeforeRightClick_EventVTableFunc (void *thisPtr,
                                                                         LPDISPATCH sh,
                                                                         LPDISPATCH target,
                                                                         VBOOL *cancel);

static void CVIFUNC AppEventsRegOnSheetActivate_EventVTableFunc (void *thisPtr,
                                                                 LPDISPATCH sh);

static void CVIFUNC AppEventsRegOnSheetDeactivate_EventVTableFunc (void *thisPtr,
                                                                   LPDISPATCH sh);

static void CVIFUNC AppEventsRegOnSheetCalculate_EventVTableFunc (void *thisPtr,
                                                                  LPDISPATCH sh);

static void CVIFUNC AppEventsRegOnSheetChange_EventVTableFunc (void *thisPtr,
                                                               LPDISPATCH sh,
                                                               LPDISPATCH target);

static void CVIFUNC AppEventsRegOnWorkbookOpen_EventVTableFunc (void *thisPtr,
                                                                LPDISPATCH wb);

static void CVIFUNC AppEventsRegOnWorkbookActivate_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH wb);

static void CVIFUNC AppEventsRegOnWorkbookDeactivate_EventVTableFunc (void *thisPtr,
                                                                      LPDISPATCH wb);

static void CVIFUNC AppEventsRegOnWorkbookBeforeClose_EventVTableFunc (void *thisPtr,
                                                                       LPDISPATCH wb,
                                                                       VBOOL *cancel);

static void CVIFUNC AppEventsRegOnWorkbookBeforeSave_EventVTableFunc (void *thisPtr,
                                                                      LPDISPATCH wb,
                                                                      VBOOL saveAsUI,
                                                                      VBOOL *cancel);

static void CVIFUNC AppEventsRegOnWorkbookBeforePrint_EventVTableFunc (void *thisPtr,
                                                                       LPDISPATCH wb,
                                                                       VBOOL *cancel);

static void CVIFUNC AppEventsRegOnWorkbookNewSheet_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH wb,
                                                                    LPDISPATCH sh);

static void CVIFUNC AppEventsRegOnWorkbookAddinInstall_EventVTableFunc (void *thisPtr,
                                                                        LPDISPATCH wb);

static void CVIFUNC AppEventsRegOnWorkbookAddinUninstall_EventVTableFunc (void *thisPtr,
                                                                          LPDISPATCH wb);

static void CVIFUNC AppEventsRegOnWindowResize_EventVTableFunc (void *thisPtr,
                                                                LPDISPATCH wb,
                                                                LPDISPATCH wn);

static void CVIFUNC AppEventsRegOnWindowActivate_EventVTableFunc (void *thisPtr,
                                                                  LPDISPATCH wb,
                                                                  LPDISPATCH wn);

static void CVIFUNC AppEventsRegOnWindowDeactivate_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH wb,
                                                                    LPDISPATCH wn);

static void CVIFUNC AppEventsRegOnSheetFollowHyperlink_EventVTableFunc (void *thisPtr,
                                                                        LPDISPATCH sh,
                                                                        LPDISPATCH target);

static void CVIFUNC ChartEventsRegOnActivate_EventVTableFunc (void *thisPtr);

static void CVIFUNC ChartEventsRegOnDeactivate_EventVTableFunc (void *thisPtr);

static void CVIFUNC ChartEventsRegOnResize_EventVTableFunc (void *thisPtr);

static void CVIFUNC ChartEventsRegOnMouseDown_EventVTableFunc (void *thisPtr,
                                                               long button,
                                                               long shift,
                                                               long x, long y);

static void CVIFUNC ChartEventsRegOnMouseUp_EventVTableFunc (void *thisPtr,
                                                             long button,
                                                             long shift, long x,
                                                             long y);

static void CVIFUNC ChartEventsRegOnMouseMove_EventVTableFunc (void *thisPtr,
                                                               long button,
                                                               long shift,
                                                               long x, long y);

static void CVIFUNC ChartEventsRegOnBeforeRightClick_EventVTableFunc (void *thisPtr,
                                                                      VBOOL *cancel);

static void CVIFUNC ChartEventsRegOnDragPlot_EventVTableFunc (void *thisPtr);

static void CVIFUNC ChartEventsRegOnDragOver_EventVTableFunc (void *thisPtr);

static void CVIFUNC ChartEventsRegOnBeforeDoubleClick_EventVTableFunc (void *thisPtr,
                                                                       long elementID,
                                                                       long arg1,
                                                                       long arg2,
                                                                       VBOOL *cancel);

static void CVIFUNC ChartEventsRegOnSelect_EventVTableFunc (void *thisPtr,
                                                            long elementID,
                                                            long arg1, long arg2);

static void CVIFUNC ChartEventsRegOnSeriesChange_EventVTableFunc (void *thisPtr,
                                                                  long seriesIndex,
                                                                  long pointIndex);

static void CVIFUNC ChartEventsRegOnCalculate_EventVTableFunc (void *thisPtr);

static void CVIFUNC DocEventsRegOnSelectionChange_EventVTableFunc (void *thisPtr,
                                                                   LPDISPATCH target);

static void CVIFUNC DocEventsRegOnBeforeDoubleClick_EventVTableFunc (void *thisPtr,
                                                                     LPDISPATCH target,
                                                                     VBOOL *cancel);

static void CVIFUNC DocEventsRegOnBeforeRightClick_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH target,
                                                                    VBOOL *cancel);

static void CVIFUNC DocEventsRegOnActivate_EventVTableFunc (void *thisPtr);

static void CVIFUNC DocEventsRegOnDeactivate_EventVTableFunc (void *thisPtr);

static void CVIFUNC DocEventsRegOnCalculate_EventVTableFunc (void *thisPtr);

static void CVIFUNC DocEventsRegOnChange_EventVTableFunc (void *thisPtr,
                                                          LPDISPATCH target);

static void CVIFUNC DocEventsRegOnFollowHyperlink_EventVTableFunc (void *thisPtr,
                                                                   LPDISPATCH target);

static void CVIFUNC WorkbkEventsRegOnOpen_EventVTableFunc (void *thisPtr);

static void CVIFUNC WorkbkEventsRegOnActivate_EventVTableFunc (void *thisPtr);

static void CVIFUNC WorkbkEventsRegOnDeactivate_EventVTableFunc (void *thisPtr);

static void CVIFUNC WorkbkEventsRegOnBeforeClose_EventVTableFunc (void *thisPtr,
                                                                  VBOOL *cancel);

static void CVIFUNC WorkbkEventsRegOnBeforeSave_EventVTableFunc (void *thisPtr,
                                                                 VBOOL saveAsUI,
                                                                 VBOOL *cancel);

static void CVIFUNC WorkbkEventsRegOnBeforePrint_EventVTableFunc (void *thisPtr,
                                                                  VBOOL *cancel);

static void CVIFUNC WorkbkEventsRegOnNewSheet_EventVTableFunc (void *thisPtr,
                                                               LPDISPATCH sh);

static void CVIFUNC WorkbkEventsRegOnAddinInstall_EventVTableFunc (void *thisPtr);

static void CVIFUNC WorkbkEventsRegOnAddinUninstall_EventVTableFunc (void *thisPtr);

static void CVIFUNC WorkbkEventsRegOnWindowResize_EventVTableFunc (void *thisPtr,
                                                                   LPDISPATCH wn);

static void CVIFUNC WorkbkEventsRegOnWindowActivate_EventVTableFunc (void *thisPtr,
                                                                     LPDISPATCH wn);

static void CVIFUNC WorkbkEventsRegOnWindowDeactivate_EventVTableFunc (void *thisPtr,
                                                                       LPDISPATCH wn);

static void CVIFUNC WorkbkEventsRegOnSheetSelectionChange_EventVTableFunc (void *thisPtr,
                                                                           LPDISPATCH sh,
                                                                           LPDISPATCH target);

static void CVIFUNC WorkbkEventsRegOnSheetBeforeDoubleClick_EventVTableFunc (void *thisPtr,
                                                                             LPDISPATCH sh,
                                                                             LPDISPATCH target,
                                                                             VBOOL *cancel);

static void CVIFUNC WorkbkEventsRegOnSheetBeforeRightClick_EventVTableFunc (void *thisPtr,
                                                                            LPDISPATCH sh,
                                                                            LPDISPATCH target,
                                                                            VBOOL *cancel);

static void CVIFUNC WorkbkEventsRegOnSheetActivate_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH sh);

static void CVIFUNC WorkbkEventsRegOnSheetDeactivate_EventVTableFunc (void *thisPtr,
                                                                      LPDISPATCH sh);

static void CVIFUNC WorkbkEventsRegOnSheetCalculate_EventVTableFunc (void *thisPtr,
                                                                     LPDISPATCH sh);

static void CVIFUNC WorkbkEventsRegOnSheetChange_EventVTableFunc (void *thisPtr,
                                                                  LPDISPATCH sh,
                                                                  LPDISPATCH target);

static void CVIFUNC WorkbkEventsRegOnSheetFollowHyperlink_EventVTableFunc (void *thisPtr,
                                                                           LPDISPATCH sh,
                                                                           LPDISPATCH target);

static void CVIFUNC OLEObjectEventsRegOnGotFocus_EventVTableFunc (void *thisPtr);

static void CVIFUNC OLEObjectEventsRegOnLostFocus_EventVTableFunc (void *thisPtr);

static CA_PARAMDATA _RefreshEvents_RegOnBeforeRefresh_CA_PARAMDATA[] =
	{
		{"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _RefreshEvents_RegOnAfterRefresh_CA_PARAMDATA[] =
	{
		{"success", VT_BOOL}
	};

static CA_METHODDATA _RefreshEvents_CA_METHODDATA[] =
	{
		{"BeforeRefresh", _RefreshEvents_RegOnBeforeRefresh_CA_PARAMDATA, 1596, 0, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"AfterRefresh", _RefreshEvents_RegOnAfterRefresh_CA_PARAMDATA, 1597, 1, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY}
	};

static CA_INTERFACEDATA _RefreshEvents_CA_INTERFACEDATA =
	{
		_RefreshEvents_CA_METHODDATA,
        (unsigned int)(sizeof (_RefreshEvents_CA_METHODDATA) / sizeof (*_RefreshEvents_CA_METHODDATA))
	};

static void * _RefreshEvents_EventVTable[] =
	{
		RefreshEventsRegOnBeforeRefresh_EventVTableFunc,
        RefreshEventsRegOnAfterRefresh_EventVTableFunc
	};

static CAEventClassDefn _RefreshEvents_CAEventClassDefn =
	{
		(int)sizeof(CAEventClassDefn),
        &Excel_IID_RefreshEvents,
        _RefreshEvents_EventVTable,
        &_RefreshEvents_CA_INTERFACEDATA,
        0
	};

static CA_PARAMDATA _AppEvents_RegOnNewWorkbook_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnSheetSelectionChange_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnSheetBeforeDoubleClick_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _AppEvents_RegOnSheetBeforeRightClick_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _AppEvents_RegOnSheetActivate_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnSheetDeactivate_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnSheetCalculate_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnSheetChange_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnWorkbookOpen_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnWorkbookActivate_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnWorkbookDeactivate_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnWorkbookBeforeClose_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _AppEvents_RegOnWorkbookBeforeSave_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH},
        {"saveAsUI", VT_BOOL},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _AppEvents_RegOnWorkbookBeforePrint_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _AppEvents_RegOnWorkbookNewSheet_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH},
        {"sh", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnWorkbookAddinInstall_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnWorkbookAddinUninstall_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnWindowResize_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH},
        {"wn", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnWindowActivate_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH},
        {"wn", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnWindowDeactivate_CA_PARAMDATA[] =
	{
		{"wb", VT_DISPATCH},
        {"wn", VT_DISPATCH}
	};

static CA_PARAMDATA _AppEvents_RegOnSheetFollowHyperlink_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH}
	};

static CA_METHODDATA _AppEvents_CA_METHODDATA[] =
	{
		{"NewWorkbook", _AppEvents_RegOnNewWorkbook_CA_PARAMDATA, 1565, 0, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"SheetSelectionChange", _AppEvents_RegOnSheetSelectionChange_CA_PARAMDATA, 1558, 1, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"SheetBeforeDoubleClick", _AppEvents_RegOnSheetBeforeDoubleClick_CA_PARAMDATA, 1559, 2, CC_STDCALL, 3, DISPATCH_METHOD, VT_EMPTY},
        {"SheetBeforeRightClick", _AppEvents_RegOnSheetBeforeRightClick_CA_PARAMDATA, 1560, 3, CC_STDCALL, 3, DISPATCH_METHOD, VT_EMPTY},
        {"SheetActivate", _AppEvents_RegOnSheetActivate_CA_PARAMDATA, 1561, 4, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"SheetDeactivate", _AppEvents_RegOnSheetDeactivate_CA_PARAMDATA, 1562, 5, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"SheetCalculate", _AppEvents_RegOnSheetCalculate_CA_PARAMDATA, 1563, 6, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"SheetChange", _AppEvents_RegOnSheetChange_CA_PARAMDATA, 1564, 7, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"WorkbookOpen", _AppEvents_RegOnWorkbookOpen_CA_PARAMDATA, 1567, 8, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"WorkbookActivate", _AppEvents_RegOnWorkbookActivate_CA_PARAMDATA, 1568, 9, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"WorkbookDeactivate", _AppEvents_RegOnWorkbookDeactivate_CA_PARAMDATA, 1569, 10, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"WorkbookBeforeClose", _AppEvents_RegOnWorkbookBeforeClose_CA_PARAMDATA, 1570, 11, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"WorkbookBeforeSave", _AppEvents_RegOnWorkbookBeforeSave_CA_PARAMDATA, 1571, 12, CC_STDCALL, 3, DISPATCH_METHOD, VT_EMPTY},
        {"WorkbookBeforePrint", _AppEvents_RegOnWorkbookBeforePrint_CA_PARAMDATA, 1572, 13, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"WorkbookNewSheet", _AppEvents_RegOnWorkbookNewSheet_CA_PARAMDATA, 1573, 14, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"WorkbookAddinInstall", _AppEvents_RegOnWorkbookAddinInstall_CA_PARAMDATA, 1574, 15, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"WorkbookAddinUninstall", _AppEvents_RegOnWorkbookAddinUninstall_CA_PARAMDATA, 1575, 16, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"WindowResize", _AppEvents_RegOnWindowResize_CA_PARAMDATA, 1554, 17, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"WindowActivate", _AppEvents_RegOnWindowActivate_CA_PARAMDATA, 1556, 18, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"WindowDeactivate", _AppEvents_RegOnWindowDeactivate_CA_PARAMDATA, 1557, 19, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"SheetFollowHyperlink", _AppEvents_RegOnSheetFollowHyperlink_CA_PARAMDATA, 1854, 20, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY}
	};

static CA_INTERFACEDATA _AppEvents_CA_INTERFACEDATA =
	{
		_AppEvents_CA_METHODDATA,
        (unsigned int)(sizeof (_AppEvents_CA_METHODDATA) / sizeof (*_AppEvents_CA_METHODDATA))
	};

static void * _AppEvents_EventVTable[] =
	{
		AppEventsRegOnNewWorkbook_EventVTableFunc,
        AppEventsRegOnSheetSelectionChange_EventVTableFunc,
        AppEventsRegOnSheetBeforeDoubleClick_EventVTableFunc,
        AppEventsRegOnSheetBeforeRightClick_EventVTableFunc,
        AppEventsRegOnSheetActivate_EventVTableFunc,
        AppEventsRegOnSheetDeactivate_EventVTableFunc,
        AppEventsRegOnSheetCalculate_EventVTableFunc,
        AppEventsRegOnSheetChange_EventVTableFunc,
        AppEventsRegOnWorkbookOpen_EventVTableFunc,
        AppEventsRegOnWorkbookActivate_EventVTableFunc,
        AppEventsRegOnWorkbookDeactivate_EventVTableFunc,
        AppEventsRegOnWorkbookBeforeClose_EventVTableFunc,
        AppEventsRegOnWorkbookBeforeSave_EventVTableFunc,
        AppEventsRegOnWorkbookBeforePrint_EventVTableFunc,
        AppEventsRegOnWorkbookNewSheet_EventVTableFunc,
        AppEventsRegOnWorkbookAddinInstall_EventVTableFunc,
        AppEventsRegOnWorkbookAddinUninstall_EventVTableFunc,
        AppEventsRegOnWindowResize_EventVTableFunc,
        AppEventsRegOnWindowActivate_EventVTableFunc,
        AppEventsRegOnWindowDeactivate_EventVTableFunc,
        AppEventsRegOnSheetFollowHyperlink_EventVTableFunc
	};

static CAEventClassDefn _AppEvents_CAEventClassDefn =
	{
		(int)sizeof(CAEventClassDefn),
        &Excel_IID_AppEvents,
        _AppEvents_EventVTable,
        &_AppEvents_CA_INTERFACEDATA,
        0
	};

static CA_PARAMDATA _ChartEvents_RegOnMouseDown_CA_PARAMDATA[] =
	{
		{"button", VT_I4},
        {"shift", VT_I4},
        {"x", VT_I4},
        {"y", VT_I4}
	};

static CA_PARAMDATA _ChartEvents_RegOnMouseUp_CA_PARAMDATA[] =
	{
		{"button", VT_I4},
        {"shift", VT_I4},
        {"x", VT_I4},
        {"y", VT_I4}
	};

static CA_PARAMDATA _ChartEvents_RegOnMouseMove_CA_PARAMDATA[] =
	{
		{"button", VT_I4},
        {"shift", VT_I4},
        {"x", VT_I4},
        {"y", VT_I4}
	};

static CA_PARAMDATA _ChartEvents_RegOnBeforeRightClick_CA_PARAMDATA[] =
	{
		{"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _ChartEvents_RegOnBeforeDoubleClick_CA_PARAMDATA[] =
	{
		{"elementID", VT_I4},
        {"arg1", VT_I4},
        {"arg2", VT_I4},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _ChartEvents_RegOnSelect_CA_PARAMDATA[] =
	{
		{"elementID", VT_I4},
        {"arg1", VT_I4},
        {"arg2", VT_I4}
	};

static CA_PARAMDATA _ChartEvents_RegOnSeriesChange_CA_PARAMDATA[] =
	{
		{"seriesIndex", VT_I4},
        {"pointIndex", VT_I4}
	};

static CA_METHODDATA _ChartEvents_CA_METHODDATA[] =
	{
		{"Activate", NULL, 304, 0, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"Deactivate", NULL, 1530, 1, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"Resize", NULL, 256, 2, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"MouseDown", _ChartEvents_RegOnMouseDown_CA_PARAMDATA, 1531, 3, CC_STDCALL, 4, DISPATCH_METHOD, VT_EMPTY},
        {"MouseUp", _ChartEvents_RegOnMouseUp_CA_PARAMDATA, 1532, 4, CC_STDCALL, 4, DISPATCH_METHOD, VT_EMPTY},
        {"MouseMove", _ChartEvents_RegOnMouseMove_CA_PARAMDATA, 1533, 5, CC_STDCALL, 4, DISPATCH_METHOD, VT_EMPTY},
        {"BeforeRightClick", _ChartEvents_RegOnBeforeRightClick_CA_PARAMDATA, 1534, 6, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"DragPlot", NULL, 1535, 7, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"DragOver", NULL, 1536, 8, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"BeforeDoubleClick", _ChartEvents_RegOnBeforeDoubleClick_CA_PARAMDATA, 1537, 9, CC_STDCALL, 4, DISPATCH_METHOD, VT_EMPTY},
        {"Select", _ChartEvents_RegOnSelect_CA_PARAMDATA, 235, 10, CC_STDCALL, 3, DISPATCH_METHOD, VT_EMPTY},
        {"SeriesChange", _ChartEvents_RegOnSeriesChange_CA_PARAMDATA, 1538, 11, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"Calculate", NULL, 279, 12, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY}
	};

static CA_INTERFACEDATA _ChartEvents_CA_INTERFACEDATA =
	{
		_ChartEvents_CA_METHODDATA,
        (unsigned int)(sizeof (_ChartEvents_CA_METHODDATA) / sizeof (*_ChartEvents_CA_METHODDATA))
	};

static void * _ChartEvents_EventVTable[] =
	{
		ChartEventsRegOnActivate_EventVTableFunc,
        ChartEventsRegOnDeactivate_EventVTableFunc,
        ChartEventsRegOnResize_EventVTableFunc,
        ChartEventsRegOnMouseDown_EventVTableFunc,
        ChartEventsRegOnMouseUp_EventVTableFunc,
        ChartEventsRegOnMouseMove_EventVTableFunc,
        ChartEventsRegOnBeforeRightClick_EventVTableFunc,
        ChartEventsRegOnDragPlot_EventVTableFunc,
        ChartEventsRegOnDragOver_EventVTableFunc,
        ChartEventsRegOnBeforeDoubleClick_EventVTableFunc,
        ChartEventsRegOnSelect_EventVTableFunc,
        ChartEventsRegOnSeriesChange_EventVTableFunc,
        ChartEventsRegOnCalculate_EventVTableFunc
	};

static CAEventClassDefn _ChartEvents_CAEventClassDefn =
	{
		(int)sizeof(CAEventClassDefn),
        &Excel_IID_ChartEvents,
        _ChartEvents_EventVTable,
        &_ChartEvents_CA_INTERFACEDATA,
        0
	};

static CA_PARAMDATA _DocEvents_RegOnSelectionChange_CA_PARAMDATA[] =
	{
		{"target", VT_DISPATCH}
	};

static CA_PARAMDATA _DocEvents_RegOnBeforeDoubleClick_CA_PARAMDATA[] =
	{
		{"target", VT_DISPATCH},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _DocEvents_RegOnBeforeRightClick_CA_PARAMDATA[] =
	{
		{"target", VT_DISPATCH},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _DocEvents_RegOnChange_CA_PARAMDATA[] =
	{
		{"target", VT_DISPATCH}
	};

static CA_PARAMDATA _DocEvents_RegOnFollowHyperlink_CA_PARAMDATA[] =
	{
		{"target", VT_DISPATCH}
	};

static CA_METHODDATA _DocEvents_CA_METHODDATA[] =
	{
		{"SelectionChange", _DocEvents_RegOnSelectionChange_CA_PARAMDATA, 1543, 0, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"BeforeDoubleClick", _DocEvents_RegOnBeforeDoubleClick_CA_PARAMDATA, 1537, 1, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"BeforeRightClick", _DocEvents_RegOnBeforeRightClick_CA_PARAMDATA, 1534, 2, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"Activate", NULL, 304, 3, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"Deactivate", NULL, 1530, 4, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"Calculate", NULL, 279, 5, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"Change", _DocEvents_RegOnChange_CA_PARAMDATA, 1545, 6, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"FollowHyperlink", _DocEvents_RegOnFollowHyperlink_CA_PARAMDATA, 1470, 7, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY}
	};

static CA_INTERFACEDATA _DocEvents_CA_INTERFACEDATA =
	{
		_DocEvents_CA_METHODDATA,
        (unsigned int)(sizeof (_DocEvents_CA_METHODDATA) / sizeof (*_DocEvents_CA_METHODDATA))
	};

static void * _DocEvents_EventVTable[] =
	{
		DocEventsRegOnSelectionChange_EventVTableFunc,
        DocEventsRegOnBeforeDoubleClick_EventVTableFunc,
        DocEventsRegOnBeforeRightClick_EventVTableFunc,
        DocEventsRegOnActivate_EventVTableFunc,
        DocEventsRegOnDeactivate_EventVTableFunc,
        DocEventsRegOnCalculate_EventVTableFunc,
        DocEventsRegOnChange_EventVTableFunc,
        DocEventsRegOnFollowHyperlink_EventVTableFunc
	};

static CAEventClassDefn _DocEvents_CAEventClassDefn =
	{
		(int)sizeof(CAEventClassDefn),
        &Excel_IID_DocEvents,
        _DocEvents_EventVTable,
        &_DocEvents_CA_INTERFACEDATA,
        0
	};

static CA_PARAMDATA _WorkbkEvents_RegOnBeforeClose_CA_PARAMDATA[] =
	{
		{"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnBeforeSave_CA_PARAMDATA[] =
	{
		{"saveAsUI", VT_BOOL},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnBeforePrint_CA_PARAMDATA[] =
	{
		{"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnNewSheet_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnWindowResize_CA_PARAMDATA[] =
	{
		{"wn", VT_DISPATCH}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnWindowActivate_CA_PARAMDATA[] =
	{
		{"wn", VT_DISPATCH}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnWindowDeactivate_CA_PARAMDATA[] =
	{
		{"wn", VT_DISPATCH}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnSheetSelectionChange_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnSheetBeforeDoubleClick_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnSheetBeforeRightClick_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH},
        {"cancel", VT_BOOL | VT_BYREF}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnSheetActivate_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnSheetDeactivate_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnSheetCalculate_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnSheetChange_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH}
	};

static CA_PARAMDATA _WorkbkEvents_RegOnSheetFollowHyperlink_CA_PARAMDATA[] =
	{
		{"sh", VT_DISPATCH},
        {"target", VT_DISPATCH}
	};

static CA_METHODDATA _WorkbkEvents_CA_METHODDATA[] =
	{
		{"Open", NULL, 682, 0, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"Activate", NULL, 304, 1, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"Deactivate", NULL, 1530, 2, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"BeforeClose", _WorkbkEvents_RegOnBeforeClose_CA_PARAMDATA, 1546, 3, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"BeforeSave", _WorkbkEvents_RegOnBeforeSave_CA_PARAMDATA, 1547, 4, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"BeforePrint", _WorkbkEvents_RegOnBeforePrint_CA_PARAMDATA, 1549, 5, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"NewSheet", _WorkbkEvents_RegOnNewSheet_CA_PARAMDATA, 1550, 6, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"AddinInstall", NULL, 1552, 7, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"AddinUninstall", NULL, 1553, 8, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"WindowResize", _WorkbkEvents_RegOnWindowResize_CA_PARAMDATA, 1554, 9, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"WindowActivate", _WorkbkEvents_RegOnWindowActivate_CA_PARAMDATA, 1556, 10, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"WindowDeactivate", _WorkbkEvents_RegOnWindowDeactivate_CA_PARAMDATA, 1557, 11, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"SheetSelectionChange", _WorkbkEvents_RegOnSheetSelectionChange_CA_PARAMDATA, 1558, 12, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"SheetBeforeDoubleClick", _WorkbkEvents_RegOnSheetBeforeDoubleClick_CA_PARAMDATA, 1559, 13, CC_STDCALL, 3, DISPATCH_METHOD, VT_EMPTY},
        {"SheetBeforeRightClick", _WorkbkEvents_RegOnSheetBeforeRightClick_CA_PARAMDATA, 1560, 14, CC_STDCALL, 3, DISPATCH_METHOD, VT_EMPTY},
        {"SheetActivate", _WorkbkEvents_RegOnSheetActivate_CA_PARAMDATA, 1561, 15, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"SheetDeactivate", _WorkbkEvents_RegOnSheetDeactivate_CA_PARAMDATA, 1562, 16, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"SheetCalculate", _WorkbkEvents_RegOnSheetCalculate_CA_PARAMDATA, 1563, 17, CC_STDCALL, 1, DISPATCH_METHOD, VT_EMPTY},
        {"SheetChange", _WorkbkEvents_RegOnSheetChange_CA_PARAMDATA, 1564, 18, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY},
        {"SheetFollowHyperlink", _WorkbkEvents_RegOnSheetFollowHyperlink_CA_PARAMDATA, 1854, 19, CC_STDCALL, 2, DISPATCH_METHOD, VT_EMPTY}
	};

static CA_INTERFACEDATA _WorkbkEvents_CA_INTERFACEDATA =
	{
		_WorkbkEvents_CA_METHODDATA,
        (unsigned int)(sizeof (_WorkbkEvents_CA_METHODDATA) / sizeof (*_WorkbkEvents_CA_METHODDATA))
	};

static void * _WorkbkEvents_EventVTable[] =
	{
		WorkbkEventsRegOnOpen_EventVTableFunc,
        WorkbkEventsRegOnActivate_EventVTableFunc,
        WorkbkEventsRegOnDeactivate_EventVTableFunc,
        WorkbkEventsRegOnBeforeClose_EventVTableFunc,
        WorkbkEventsRegOnBeforeSave_EventVTableFunc,
        WorkbkEventsRegOnBeforePrint_EventVTableFunc,
        WorkbkEventsRegOnNewSheet_EventVTableFunc,
        WorkbkEventsRegOnAddinInstall_EventVTableFunc,
        WorkbkEventsRegOnAddinUninstall_EventVTableFunc,
        WorkbkEventsRegOnWindowResize_EventVTableFunc,
        WorkbkEventsRegOnWindowActivate_EventVTableFunc,
        WorkbkEventsRegOnWindowDeactivate_EventVTableFunc,
        WorkbkEventsRegOnSheetSelectionChange_EventVTableFunc,
        WorkbkEventsRegOnSheetBeforeDoubleClick_EventVTableFunc,
        WorkbkEventsRegOnSheetBeforeRightClick_EventVTableFunc,
        WorkbkEventsRegOnSheetActivate_EventVTableFunc,
        WorkbkEventsRegOnSheetDeactivate_EventVTableFunc,
        WorkbkEventsRegOnSheetCalculate_EventVTableFunc,
        WorkbkEventsRegOnSheetChange_EventVTableFunc,
        WorkbkEventsRegOnSheetFollowHyperlink_EventVTableFunc
	};

static CAEventClassDefn _WorkbkEvents_CAEventClassDefn =
	{
		(int)sizeof(CAEventClassDefn),
        &Excel_IID_WorkbkEvents,
        _WorkbkEvents_EventVTable,
        &_WorkbkEvents_CA_INTERFACEDATA,
        0
	};

static CA_METHODDATA _OLEObjectEvents_CA_METHODDATA[] =
	{
		{"GotFocus", NULL, 1541, 0, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY},
        {"LostFocus", NULL, 1542, 1, CC_STDCALL, 0, DISPATCH_METHOD, VT_EMPTY}
	};

static CA_INTERFACEDATA _OLEObjectEvents_CA_INTERFACEDATA =
	{
		_OLEObjectEvents_CA_METHODDATA,
        (unsigned int)(sizeof (_OLEObjectEvents_CA_METHODDATA) / sizeof (*_OLEObjectEvents_CA_METHODDATA))
	};

static void * _OLEObjectEvents_EventVTable[] =
	{
		OLEObjectEventsRegOnGotFocus_EventVTableFunc,
        OLEObjectEventsRegOnLostFocus_EventVTableFunc
	};

static CAEventClassDefn _OLEObjectEvents_CAEventClassDefn =
	{
		(int)sizeof(CAEventClassDefn),
        &Excel_IID_OLEObjectEvents,
        _OLEObjectEvents_EventVTable,
        &_OLEObjectEvents_CA_INTERFACEDATA,
        0
	};

const IID Excel_IID_QueryTable =
	{
		0x24428, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_App =
	{
		0x208D5, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Chart =
	{
		0x208D6, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Worksheet =
	{
		0x208D8, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Workbook =
	{
		0x208DA, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_OLEObject =
	{
		0x208A2, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_RefreshEvents =
	{
		0x2441B, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_AppEvents =
	{
		0x24413, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ChartEvents =
	{
		0x2440F, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_DocEvents =
	{
		0x24411, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_WorkbkEvents =
	{
		0x24412, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_OLEObjectEvents =
	{
		0x24410, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Adjustments =
	{
		0xC0310, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_CalloutFormat =
	{
		0xC0311, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ColorFormat =
	{
		0xC0312, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_LineFormat =
	{
		0xC0317, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ShapeNode =
	{
		0xC0318, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ShapeNodes =
	{
		0xC0319, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PictureFmt =
	{
		0xC031A, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ShadowFormat =
	{
		0xC031B, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_TextEffectFmt =
	{
		0xC031F, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ThreeDFmt =
	{
		0xC0321, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_FillFormat =
	{
		0xC0314, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Sheets =
	{
		0x208D7, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Workbooks =
	{
		0x208DB, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PublishObject =
	{
		0x24444, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_DefaultWebOptions =
	{
		0x24448, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_WebOptions =
	{
		0x24449, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_TreeviewControl =
	{
		0x2444B, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_CubeField =
	{
		0x2444C, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_CubeFields =
	{
		0x2444D, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Font =
	{
		0x2084D, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Window =
	{
		0x20893, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Windows =
	{
		0x20892, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_WorksheetFunction =
	{
		0x20845, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Range =
	{
		0x20846, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_VPageBreak =
	{
		0x24402, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_HPageBreak =
	{
		0x24401, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_HPageBreaks =
	{
		0x24404, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_VPageBreaks =
	{
		0x24405, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_RecentFile =
	{
		0x24407, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_RecentFiles =
	{
		0x24406, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Style =
	{
		0x20852, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Styles =
	{
		0x20853, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Borders =
	{
		0x20855, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_AddIn =
	{
		0x20857, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_AddIns =
	{
		0x20858, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Areas =
	{
		0x20860, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Charts =
	{
		0x2086D, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotCache =
	{
		0x2441C, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotCaches =
	{
		0x2441D, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotFormula =
	{
		0x2441E, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotFormulas =
	{
		0x2441F, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotTable =
	{
		0x20872, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotTables =
	{
		0x20873, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotField =
	{
		0x20874, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotFields =
	{
		0x20875, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_CalculatedFields =
	{
		0x24420, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotItem =
	{
		0x20876, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotItems =
	{
		0x20877, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_CalculatedItems =
	{
		0x24421, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Characters =
	{
		0x20878, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Dialogs =
	{
		0x20879, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Dialog =
	{
		0x2087A, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_SoundNote =
	{
		0x2087B, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Panes =
	{
		0x20894, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Pane =
	{
		0x20895, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Scenarios =
	{
		0x20896, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Scenario =
	{
		0x20897, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_OLEObjects =
	{
		0x208A3, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_RoutingSlip =
	{
		0x208AA, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Outline =
	{
		0x208AB, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Worksheets =
	{
		0x208B1, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PageSetup =
	{
		0x208B4, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Names =
	{
		0x208B8, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Name =
	{
		0x208B9, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ChartObject =
	{
		0x208CF, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ChartObjects =
	{
		0x208D0, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Mailer =
	{
		0x208D1, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_CustomViews =
	{
		0x24422, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_CustomView =
	{
		0x24423, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_FormatConditions =
	{
		0x24424, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_FormatCondition =
	{
		0x24425, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Comments =
	{
		0x24426, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Comment =
	{
		0x24427, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_QueryTables =
	{
		0x24429, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Parameter =
	{
		0x2442A, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Parameters =
	{
		0x2442B, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ODBCError =
	{
		0x2442C, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ODBCErrors =
	{
		0x2442D, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Validation =
	{
		0x2442F, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Hyperlinks =
	{
		0x24430, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Hyperlink =
	{
		0x24431, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_AutoFilter =
	{
		0x24432, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Filters =
	{
		0x24433, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Filter =
	{
		0x24434, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_AutoCorrect =
	{
		0x208D4, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Border =
	{
		0x20854, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Interior =
	{
		0x20870, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ChartFillFormat =
	{
		0x24435, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ChartColorFormat =
	{
		0x24436, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Axis =
	{
		0x20848, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ChartTitle =
	{
		0x20849, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_AxisTitle =
	{
		0x2084A, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ChartGroup =
	{
		0x20859, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ChartGroups =
	{
		0x2085A, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Axes =
	{
		0x2085B, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Points =
	{
		0x20869, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Point =
	{
		0x2086A, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Series =
	{
		0x2086B, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_SeriesCollection =
	{
		0x2086C, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_DataLabel =
	{
		0x208B2, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_DataLabels =
	{
		0x208B3, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_LegendEntry =
	{
		0x208BA, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_LegendEntries =
	{
		0x208BB, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_LegendKey =
	{
		0x208BC, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Trendlines =
	{
		0x208BD, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Trendline =
	{
		0x208BE, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Corners =
	{
		0x208C0, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_SeriesLines =
	{
		0x208C1, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_HiLoLines =
	{
		0x208C2, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Gridlines =
	{
		0x208C3, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_DropLines =
	{
		0x208C4, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_LeaderLines =
	{
		0x24437, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_UpBars =
	{
		0x208C5, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_DownBars =
	{
		0x208C6, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Floor =
	{
		0x208C7, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Walls =
	{
		0x208C8, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_TickLabels =
	{
		0x208C9, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PlotArea =
	{
		0x208CB, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ChartArea =
	{
		0x208CC, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Legend =
	{
		0x208CD, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ErrorBars =
	{
		0x208CE, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_DataTable =
	{
		0x20843, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Phonetic =
	{
		0x24438, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Shape =
	{
		0x24439, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Shapes =
	{
		0x2443A, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ShapeRange =
	{
		0x2443B, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_GroupShapes =
	{
		0x2443C, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_TextFrame =
	{
		0x2443D, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ConnectorFormat =
	{
		0x2443E, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_FreeformBuilder =
	{
		0x2443F, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_ControlFormat =
	{
		0x24440, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_OLEFormat =
	{
		0x24441, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_LinkFormat =
	{
		0x24442, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PublishObjects =
	{
		0x24443, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_OLEDBError =
	{
		0x24445, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_OLEDBErrors =
	{
		0x24446, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_Phonetics =
	{
		0x24447, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_PivotLayout =
	{
		0x2444A, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

const IID Excel_IID_DisplayUnitLabel =
	{
		0x2084C, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46
	};

HRESULT CVIFUNC Excel_GetProperty (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, unsigned int propertyID,
                                   unsigned int propertyType, void *value)
{
	HRESULT __result;

	__result = CA_PropertyGetEx (objectHandle, errorInfo, 0, propertyID,
	                             propertyType, value);

	return __result;
}

HRESULT CVIFUNC_C Excel_SetProperty (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     unsigned int propertyID,
                                     unsigned int propertyType, ...)
{
	HRESULT __result;
	va_list args;

	va_start (args, propertyType);

	__result = CA_PropertySetVEx (objectHandle, errorInfo, 0, propertyID,
	                              propertyType, args);

	va_end (args);

	return __result;
}

HRESULT CVIFUNC_C Excel_SetPropertyByRef (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          unsigned int propertyID,
                                          unsigned int propertyType, ...)
{
	HRESULT __result;
	va_list args;

	va_start (args, propertyType);

	__result = CA_PropertySetByRefVEx (objectHandle, errorInfo, 0, propertyID,
	                                   propertyType, args);

	va_end (args);

	return __result;
}

HRESULT CVIFUNC Excel_NewQueryTable (const char *server,
                                     int supportMultithreading, LCID locale,
                                     int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x59191DA1, 0xEA47, 0x11CE, 0xA5, 0x1F, 0x0, 0xAA, 0x0, 0x61,
	              0x50, 0x7F};

	__result = CA_CreateObjectByClassIdEx (&clsid, server,
	                                       &Excel_IID_QueryTable,
	                                       supportMultithreading, locale,
	                                       reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_OpenQueryTable (const char *fileName, const char *server,
                                      int supportMultithreading, LCID locale,
                                      int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x59191DA1, 0xEA47, 0x11CE, 0xA5, 0x1F, 0x0, 0xAA, 0x0, 0x61,
	              0x50, 0x7F};

	__result = CA_LoadObjectFromFileByClassIdEx (fileName, &clsid, server,
	                                             &Excel_IID_QueryTable,
	                                             supportMultithreading, locale,
	                                             reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_ActiveQueryTable (const char *server,
                                        int supportMultithreading, LCID locale,
                                        int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x59191DA1, 0xEA47, 0x11CE, 0xA5, 0x1F, 0x0, 0xAA, 0x0, 0x61,
	              0x50, 0x7F};

	__result = CA_GetActiveObjectByClassIdEx (&clsid, server,
	                                          &Excel_IID_QueryTable,
	                                          supportMultithreading, locale,
	                                          reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_QueryTableCancelRefresh (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_QueryTable, 0x635, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_QueryTableDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_QueryTable, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_QueryTableRefresh (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT backgroundQuery,
                                         VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_QueryTable, 0x589, CAVT_BOOL,
	                              returnValue, 1, __paramTypes,
	                              backgroundQuery);

	return __result;
}

HRESULT CVIFUNC Excel_QueryTableResetTimer (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_QueryTable, 0x72A, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_NewApp (const char *server, int supportMultithreading,
                              LCID locale, int reserved,
                              CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x24500, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_CreateObjectByClassIdEx (&clsid, server, &Excel_IID_App,
	                                       supportMultithreading, locale,
	                                       reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_OpenApp (const char *fileName, const char *server,
                               int supportMultithreading, LCID locale,
                               int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x24500, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_LoadObjectFromFileByClassIdEx (fileName, &clsid, server,
	                                             &Excel_IID_App,
	                                             supportMultithreading, locale,
	                                             reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_ActiveApp (const char *server, int supportMultithreading,
                                 LCID locale, int reserved,
                                 CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x24500, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_GetActiveObjectByClassIdEx (&clsid, server, &Excel_IID_App,
	                                          supportMultithreading, locale,
	                                          reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_AppCalculate (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x117, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDDEExecute (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long channel,
                                     const char *string)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x14D, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              channel, string);

	return __result;
}

HRESULT CVIFUNC Excel_AppDDEInitiate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, const char *app,
                                      const char *topic, long *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x14E, CAVT_LONG, returnValue, 2,
	                              __paramTypes, app, topic);

	return __result;
}

HRESULT CVIFUNC Excel_AppDDEPoke (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  long channel, VARIANT item, VARIANT data)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x14F, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              channel, item, data);

	return __result;
}

HRESULT CVIFUNC Excel_AppDDERequest (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long channel,
                                     const char *item, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x150, CAVT_VARIANT, returnValue, 2,
	                              __paramTypes, channel, item);

	return __result;
}

HRESULT CVIFUNC Excel_AppDDETerminate (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long channel)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x151, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              channel);

	return __result;
}

HRESULT CVIFUNC Excel_AppEvaluate (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT name,
                                   VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App, 0x1,
	                              CAVT_VARIANT, returnValue, 1, __paramTypes,
	                              name);

	return __result;
}

HRESULT CVIFUNC Excel_App_Evaluate (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT name,
                                    VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0xFFFFFFFB, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, name);

	return __result;
}

HRESULT CVIFUNC Excel_AppExecuteExcel4Macro (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *string,
                                             VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x15E, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, string);

	return __result;
}

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
                                    ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_OBJHANDLE,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x2FE, CAVT_OBJHANDLE, returnValue, 30,
	                              __paramTypes, arg1, arg2, arg3, arg4, arg5,
	                              arg6, arg7, arg8, arg9, arg10, arg11, arg12,
	                              arg13, arg14, arg15, arg16, arg17, arg18,
	                              arg19, arg20, arg21, arg22, arg23, arg24,
	                              arg25, arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_AppRange (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT cell1, VARIANT cell2,
                                ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0xC5, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, cell1, cell2);

	return __result;
}

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
                              VARIANT arg30, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x103, CAVT_VARIANT, returnValue, 31,
	                              __paramTypes, macro, arg1, arg2, arg3, arg4,
	                              arg5, arg6, arg7, arg8, arg9, arg10, arg11,
	                              arg12, arg13, arg14, arg15, arg16, arg17,
	                              arg18, arg19, arg20, arg21, arg22, arg23,
	                              arg24, arg25, arg26, arg27, arg28, arg29,
	                              arg30);

	return __result;
}

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
                                VARIANT arg30, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x326, CAVT_VARIANT, returnValue, 31,
	                              __paramTypes, macro, arg1, arg2, arg3, arg4,
	                              arg5, arg6, arg7, arg8, arg9, arg10, arg11,
	                              arg12, arg13, arg14, arg15, arg16, arg17,
	                              arg18, arg19, arg20, arg21, arg22, arg23,
	                              arg24, arg25, arg26, arg27, arg28, arg29,
	                              arg30);

	return __result;
}

HRESULT CVIFUNC Excel_AppSendKeys (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT keys,
                                   VARIANT wait)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x17F, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              keys, wait);

	return __result;
}

HRESULT CVIFUNC Excel_AppShortcutMenus (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, long index,
                                        ExcelObj_Menu *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x308, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

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
                                VARIANT arg30, ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_OBJHANDLE,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x30B, CAVT_OBJHANDLE, returnValue, 30,
	                              __paramTypes, arg1, arg2, arg3, arg4, arg5,
	                              arg6, arg7, arg8, arg9, arg10, arg11, arg12,
	                              arg13, arg14, arg15, arg16, arg17, arg18,
	                              arg19, arg20, arg21, arg22, arg23, arg24,
	                              arg25, arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_AppActivateMicrosoftApp (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               enum ExcelEnum_XlMSApplication index)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x447, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_AppAddChartAutoFormat (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT chart,
                                             const char *name,
                                             VARIANT description)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0xD8, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              chart, name, description);

	return __result;
}

HRESULT CVIFUNC Excel_AppAddCustomList (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT listArray,
                                        VARIANT byRow)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x30C, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              listArray, byRow);

	return __result;
}

HRESULT CVIFUNC Excel_AppCaller (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT index, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x13D, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AppCentimetersToPoints (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              double centimeters,
                                              double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x43E, CAVT_DOUBLE, returnValue, 1,
	                              __paramTypes, centimeters);

	return __result;
}

HRESULT CVIFUNC Excel_AppCheckSpelling (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, const char *word,
                                        VARIANT customDictionary,
                                        VARIANT ignoreUppercase,
                                        VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x1F9, CAVT_BOOL, returnValue, 3,
	                              __paramTypes, word, customDictionary,
	                              ignoreUppercase);

	return __result;
}

HRESULT CVIFUNC Excel_AppClipboardFormats (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x141, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AppConvertFormula (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT formula,
                                         enum ExcelEnum_XlReferenceStyle fromReferenceStyle,
                                         VARIANT toReferenceStyle,
                                         VARIANT toAbsolute, VARIANT relativeTo,
                                         VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x145, CAVT_VARIANT, returnValue, 5,
	                              __paramTypes, formula, fromReferenceStyle,
	                              toReferenceStyle, toAbsolute, relativeTo);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy1 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6F6, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy2 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6F7, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy3 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6F8, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy4 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6F9, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy5 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6FA, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy6 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6FB, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy7 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6FC, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy8 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6FD, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy9 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6FE, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy10 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6FF, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy11 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x700, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppDeleteChartAutoFormat (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                const char *name)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0xD9, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              name);

	return __result;
}

HRESULT CVIFUNC Excel_AppDeleteCustomList (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long listNum)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x30F, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              listNum);

	return __result;
}

HRESULT CVIFUNC Excel_AppDoubleClick (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x15D, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppFileConverters (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index1,
                                         VARIANT index2, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x3A3, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, index1, index2);

	return __result;
}

HRESULT CVIFUNC Excel_App_FindFile (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x42C, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppGetCustomListContents (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                long listNum,
                                                VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x312, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, listNum);

	return __result;
}

HRESULT CVIFUNC Excel_AppGetCustomListNum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT listArray, long *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x311, CAVT_LONG, returnValue, 1,
	                              __paramTypes, listArray);

	return __result;
}

HRESULT CVIFUNC Excel_AppGetOpenFilename (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT fileFilter,
                                          VARIANT filterIndex, VARIANT title,
                                          VARIANT buttonText,
                                          VARIANT multiSelect,
                                          VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x433, CAVT_VARIANT, returnValue, 5,
	                              __paramTypes, fileFilter, filterIndex, title,
	                              buttonText, multiSelect);

	return __result;
}

HRESULT CVIFUNC Excel_AppGetSaveAsFilename (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT initialFilename,
                                            VARIANT fileFilter,
                                            VARIANT filterIndex, VARIANT title,
                                            VARIANT buttonText,
                                            VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x434, CAVT_VARIANT, returnValue, 5,
	                              __paramTypes, initialFilename, fileFilter,
	                              filterIndex, title, buttonText);

	return __result;
}

HRESULT CVIFUNC Excel_AppGoto (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                               VARIANT reference, VARIANT scroll)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x1DB, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              reference, scroll);

	return __result;
}

HRESULT CVIFUNC Excel_AppHelp (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                               VARIANT helpFile, VARIANT helpContextID)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x162, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              helpFile, helpContextID);

	return __result;
}

HRESULT CVIFUNC Excel_AppInchesToPoints (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, double inches,
                                         double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x43F, CAVT_DOUBLE, returnValue, 1,
	                              __paramTypes, inches);

	return __result;
}

HRESULT CVIFUNC Excel_AppInputBox (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, const char *prompt,
                                   VARIANT title, VARIANT _default, VARIANT left,
                                   VARIANT top, VARIANT helpFile,
                                   VARIANT helpContextID, VARIANT type,
                                   VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x165, CAVT_VARIANT, returnValue, 8,
	                              __paramTypes, prompt, title, _default, left,
	                              top, helpFile, helpContextID, type);

	return __result;
}

HRESULT CVIFUNC Excel_AppInternational (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x16A, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AppMacroOptions (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT macro,
                                       VARIANT description, VARIANT hasMenu,
                                       VARIANT menuText, VARIANT hasShortcutKey,
                                       VARIANT shortcutKey, VARIANT category,
                                       VARIANT statusBar, VARIANT helpContextID,
                                       VARIANT helpFile)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x46F, CAVT_EMPTY, NULL, 10, __paramTypes,
	                              macro, description, hasMenu, menuText,
	                              hasShortcutKey, shortcutKey, category,
	                              statusBar, helpContextID, helpFile);

	return __result;
}

HRESULT CVIFUNC Excel_AppMailLogoff (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x3B1, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppMailLogon (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT name,
                                    VARIANT password, VARIANT downloadNewMail)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x3AF, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              name, password, downloadNewMail);

	return __result;
}

HRESULT CVIFUNC Excel_AppNextLetter (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     ExcelObj_Workbook *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x3CC, CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppOnKey (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                const char *key, VARIANT procedure)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x272, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              key, procedure);

	return __result;
}

HRESULT CVIFUNC Excel_AppOnRepeat (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, const char *text,
                                   const char *procedure)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x301, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              text, procedure);

	return __result;
}

HRESULT CVIFUNC Excel_AppOnTime (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT earliestTime, const char *procedure,
                                 VARIANT latestTime, VARIANT schedule)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_CSTRING, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x270, CAVT_EMPTY, NULL, 4, __paramTypes,
	                              earliestTime, procedure, latestTime,
	                              schedule);

	return __result;
}

HRESULT CVIFUNC Excel_AppOnUndo (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 const char *text, const char *procedure)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x302, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              text, procedure);

	return __result;
}

HRESULT CVIFUNC Excel_AppPreviousSelections (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x17A, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AppQuit (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x12E, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppRecordMacro (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT basicCode,
                                      VARIANT xlmCode)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x305, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              basicCode, xlmCode);

	return __result;
}

HRESULT CVIFUNC Excel_AppRegisteredFunctions (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index1, VARIANT index2,
                                              VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x307, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, index1, index2);

	return __result;
}

HRESULT CVIFUNC Excel_AppRegisterXLL (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, const char *filename,
                                      VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x1E, CAVT_BOOL, returnValue, 1,
	                              __paramTypes, filename);

	return __result;
}

HRESULT CVIFUNC Excel_AppRepeat (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x12D, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppResetTipWizard (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x3A0, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppSave (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                               VARIANT filename)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x11B, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              filename);

	return __result;
}

HRESULT CVIFUNC Excel_AppSaveWorkspace (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT filename)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0xD4, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              filename);

	return __result;
}

HRESULT CVIFUNC Excel_AppSetDefaultChart (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT formatName, VARIANT gallery)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0xDB, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              formatName, gallery);

	return __result;
}

HRESULT CVIFUNC Excel_AppUndo (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x12F, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppVolatile (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT _volatile)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x314, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              _volatile);

	return __result;
}

HRESULT CVIFUNC Excel_App_Wait (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT time)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x189, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              time);

	return __result;
}

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
                                      VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0xA9, CAVT_VARIANT, returnValue, 30,
	                              __paramTypes, arg1, arg2, arg3, arg4, arg5,
	                              arg6, arg7, arg8, arg9, arg10, arg11, arg12,
	                              arg13, arg14, arg15, arg16, arg17, arg18,
	                              arg19, arg20, arg21, arg22, arg23, arg24,
	                              arg25, arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_AppWait (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                               VARIANT time, VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6EA, CAVT_BOOL, returnValue, 1,
	                              __paramTypes, time);

	return __result;
}

HRESULT CVIFUNC Excel_AppGetPhonetic (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT text,
                                      char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x703, CAVT_CSTRING, returnValue, 1,
	                              __paramTypes, text);

	return __result;
}

HRESULT CVIFUNC Excel_AppDummy12 (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x70B, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppCalculateFull (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x70D, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AppFindFile (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VBOOL *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_App,
	                              0x6EB, CAVT_BOOL, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_NewChart (const char *server, int supportMultithreading,
                                LCID locale, int reserved,
                                CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20821, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_CreateObjectByClassIdEx (&clsid, server, &Excel_IID_Chart,
	                                       supportMultithreading, locale,
	                                       reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_OpenChart (const char *fileName, const char *server,
                                 int supportMultithreading, LCID locale,
                                 int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20821, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_LoadObjectFromFileByClassIdEx (fileName, &clsid, server,
	                                             &Excel_IID_Chart,
	                                             supportMultithreading, locale,
	                                             reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_ActiveChart (const char *server, int supportMultithreading,
                                   LCID locale, int reserved,
                                   CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20821, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_GetActiveObjectByClassIdEx (&clsid, server, &Excel_IID_Chart,
	                                          supportMultithreading, locale,
	                                          reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_ChartActivate (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x130, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT before, VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x227, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              before, after);

	return __result;
}

HRESULT CVIFUNC Excel_ChartDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x75, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartMove (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT before, VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x27D, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              before, after);

	return __result;
}

HRESULT CVIFUNC Excel_Chart_PrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x389, CAVT_EMPTY, NULL, 7, __paramTypes,
	                              from, to, copies, preview, activePrinter,
	                              printToFile, collate);

	return __result;
}

HRESULT CVIFUNC Excel_ChartPrintPreview (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT enableChanges)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x119, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              enableChanges);

	return __result;
}

HRESULT CVIFUNC Excel_ChartProtect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT password,
                                    VARIANT drawingObjects, VARIANT contents,
                                    VARIANT scenarios, VARIANT userInterfaceOnly)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x11A, CAVT_EMPTY, NULL, 5, __paramTypes,
	                              password, drawingObjects, contents,
	                              scenarios, userInterfaceOnly);

	return __result;
}

HRESULT CVIFUNC Excel_Chart_Dummy23 (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x10017, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartSaveAs (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, const char *filename,
                                   VARIANT fileFormat, VARIANT password,
                                   VARIANT writeResPassword,
                                   VARIANT readOnlyRecommended,
                                   VARIANT createBackup, VARIANT addToMru,
                                   VARIANT textCodepage,
                                   VARIANT textVisualLayout)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x11C, CAVT_EMPTY, NULL, 9, __paramTypes,
	                              filename, fileFormat, password,
	                              writeResPassword, readOnlyRecommended,
	                              createBackup, addToMru, textCodepage,
	                              textVisualLayout);

	return __result;
}

HRESULT CVIFUNC Excel_ChartSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT replace)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xEB, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              replace);

	return __result;
}

HRESULT CVIFUNC Excel_ChartUnprotect (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT password)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x11D, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              password);

	return __result;
}

HRESULT CVIFUNC Excel_ChartApplyDataLabels (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlDataLabelsType type,
                                            VARIANT legendKey, VARIANT autoText,
                                            VARIANT hasLeaderLines)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x97, CAVT_EMPTY, NULL, 4, __paramTypes,
	                              type, legendKey, autoText, hasLeaderLines);

	return __result;
}

HRESULT CVIFUNC Excel_ChartArcs (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT index, CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x2F8, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartAreaGroups (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x9, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartAutoFormat (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long gallery,
                                       VARIANT format)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x72, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              gallery, format);

	return __result;
}

HRESULT CVIFUNC Excel_ChartAxes (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT type,
                                 enum ExcelEnum_XlAxisGroup axisGroup,
                                 CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x17, CAVT_OBJHANDLE, returnValue, 2,
	                              __paramTypes, type, axisGroup);

	return __result;
}

HRESULT CVIFUNC Excel_ChartSetBackgroundPicture (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 const char *filename)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x4A4, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              filename);

	return __result;
}

HRESULT CVIFUNC Excel_ChartBarGroups (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xA, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartButtons (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT index,
                                    CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x22D, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartChartGroups (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x8, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartChartObjects (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x424, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartChartWizard (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT source,
                                        VARIANT gallery, VARIANT format,
                                        VARIANT plotBy, VARIANT categoryLabels,
                                        VARIANT seriesLabels, VARIANT hasLegend,
                                        VARIANT title, VARIANT categoryTitle,
                                        VARIANT valueTitle, VARIANT extraTitle)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xC4, CAVT_EMPTY, NULL, 11, __paramTypes,
	                              source, gallery, format, plotBy,
	                              categoryLabels, seriesLabels, hasLegend,
	                              title, categoryTitle, valueTitle, extraTitle);

	return __result;
}

HRESULT CVIFUNC Excel_ChartCheckBoxes (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x338, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartCheckSpelling (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT customDictionary,
                                          VARIANT ignoreUppercase,
                                          VARIANT alwaysSuggest,
                                          VARIANT spellLang)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x1F9, CAVT_EMPTY, NULL, 4, __paramTypes,
	                              customDictionary, ignoreUppercase,
	                              alwaysSuggest, spellLang);

	return __result;
}

HRESULT CVIFUNC Excel_ChartColumnGroups (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xB, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartCopyPicture (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_XlPictureAppearance appearance,
                                        enum ExcelEnum_XlCopyPictureFormat format,
                                        enum ExcelEnum_XlPictureAppearance size)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xD5, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              appearance, format, size);

	return __result;
}

HRESULT CVIFUNC Excel_ChartCreatePublisher (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT edition,
                                            enum ExcelEnum_XlPictureAppearance appearance,
                                            enum ExcelEnum_XlPictureAppearance size,
                                            VARIANT containsPICT,
                                            VARIANT containsBIFF,
                                            VARIANT containsRTF,
                                            VARIANT containsVALU)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x1CA, CAVT_EMPTY, NULL, 7, __paramTypes,
	                              edition, appearance, size, containsPICT,
	                              containsBIFF, containsRTF, containsVALU);

	return __result;
}

HRESULT CVIFUNC Excel_ChartDeselect (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x460, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartDoughnutGroups (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xE, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartDrawings (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x304, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartDrawingObjects (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x58, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartDropDowns (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x344, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEvaluate (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT name,
                                     VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x1, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, name);

	return __result;
}

HRESULT CVIFUNC Excel_Chart_Evaluate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT name,
                                      VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xFFFFFFFB, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, name);

	return __result;
}

HRESULT CVIFUNC Excel_ChartGroupBoxes (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x342, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartGroupObjects (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x459, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartGetHasAxis (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index1,
                                       VARIANT index2, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x34, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, index1, index2);

	return __result;
}

HRESULT CVIFUNC Excel_ChartSetHasAxis (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index1,
                                       VARIANT index2, VARIANT newValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x34, DISPATCH_PROPERTYPUT, CAVT_EMPTY, NULL,
	                              3, __paramTypes, index1, index2, newValue);

	return __result;
}

HRESULT CVIFUNC Excel_ChartLabels (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT index,
                                   CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x349, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartLineGroups (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xC, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartLines (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x2FF, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartListBoxes (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x340, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartLocation (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     enum ExcelEnum_XlChartLocation where,
                                     VARIANT name, ExcelObj_Chart *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x575, CAVT_OBJHANDLE, returnValue, 2,
	                              __paramTypes, where, name);

	return __result;
}

HRESULT CVIFUNC Excel_ChartOLEObjects (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x31F, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartOptionButtons (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x33A, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartOvals (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x321, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartPaste (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT type)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xD3, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              type);

	return __result;
}

HRESULT CVIFUNC Excel_ChartPictures (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x303, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartPieGroups (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xD, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartRadarGroups (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0xF, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartRectangles (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x306, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartScrollBars (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x33E, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartSeriesCollection (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x44, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartSpinners (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x346, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartTextBoxes (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x309, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartApplyCustomType (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlChartType chartType,
                                            VARIANT typeName)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x579, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              chartType, typeName);

	return __result;
}

HRESULT CVIFUNC Excel_ChartXYGroups (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x10, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartCopyChartBuild (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x57C, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartGetChartElement (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, long x, long y,
                                            long *elementID, long *arg1,
                                            long *arg2)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG,
	                               CAVT_LONG | CAVT_BYREFI,
	                               CAVT_LONG | CAVT_BYREFI,
	                               CAVT_LONG | CAVT_BYREFI};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x581, CAVT_EMPTY, NULL, 5, __paramTypes, x,
	                              y, elementID, arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_ChartSetSourceData (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          ExcelObj_Range source, VARIANT plotBy)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x585, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              source, plotBy);

	return __result;
}

HRESULT CVIFUNC Excel_ChartExport (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, const char *filename,
                                   VARIANT filterName, VARIANT interactive,
                                   VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x586, CAVT_BOOL, returnValue, 3,
	                              __paramTypes, filename, filterName,
	                              interactive);

	return __result;
}

HRESULT CVIFUNC Excel_ChartRefresh (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x589, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartPrintOut (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT from,
                                     VARIANT to, VARIANT copies, VARIANT preview,
                                     VARIANT activePrinter, VARIANT printToFile,
                                     VARIANT collate, VARIANT prToFileName)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Chart,
	                              0x6EC, CAVT_EMPTY, NULL, 8, __paramTypes,
	                              from, to, copies, preview, activePrinter,
	                              printToFile, collate, prToFileName);

	return __result;
}

HRESULT CVIFUNC Excel_NewWorksheet (const char *server,
                                    int supportMultithreading, LCID locale,
                                    int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20820, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_CreateObjectByClassIdEx (&clsid, server, &IID_IDispatch,
	                                       supportMultithreading, locale,
	                                       reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_OpenWorksheet (const char *fileName, const char *server,
                                     int supportMultithreading, LCID locale,
                                     int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20820, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_LoadObjectFromFileByClassIdEx (fileName, &clsid, server,
	                                             &IID_IDispatch,
	                                             supportMultithreading, locale,
	                                             reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_ActiveWorksheet (const char *server,
                                       int supportMultithreading, LCID locale,
                                       int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20820, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_GetActiveObjectByClassIdEx (&clsid, server, &IID_IDispatch,
	                                          supportMultithreading, locale,
	                                          reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetActivate (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x130, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetCopy (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT before,
                                     VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x227, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, before, after);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x75, CAVT_EMPTY, NULL,
	                              0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetMove (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT before,
                                     VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x27D, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, before, after);

	return __result;
}

HRESULT CVIFUNC Excel_Worksheet_PrintOut (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT from,
                                          VARIANT to, VARIANT copies,
                                          VARIANT preview, VARIANT activePrinter,
                                          VARIANT printToFile, VARIANT collate)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x389, CAVT_EMPTY,
	                              NULL, 7, __paramTypes, from, to, copies,
	                              preview, activePrinter, printToFile, collate);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetPrintPreview (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT enableChanges)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x119, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, enableChanges);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetProtect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT password,
                                        VARIANT drawingObjects, VARIANT contents,
                                        VARIANT scenarios,
                                        VARIANT userInterfaceOnly)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x11A, CAVT_EMPTY,
	                              NULL, 5, __paramTypes, password,
	                              drawingObjects, contents, scenarios,
	                              userInterfaceOnly);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetSaveAs (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       const char *filename, VARIANT fileFormat,
                                       VARIANT password,
                                       VARIANT writeResPassword,
                                       VARIANT readOnlyRecommended,
                                       VARIANT createBackup, VARIANT addToMru,
                                       VARIANT textCodepage,
                                       VARIANT textVisualLayout)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x11C, CAVT_EMPTY,
	                              NULL, 9, __paramTypes, filename, fileFormat,
	                              password, writeResPassword,
	                              readOnlyRecommended, createBackup, addToMru,
	                              textCodepage, textVisualLayout);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT replace)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0xEB, CAVT_EMPTY, NULL,
	                              1, __paramTypes, replace);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetUnprotect (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT password)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x11D, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, password);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetArcs (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x2F8, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetSetBackgroundPicture (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     const char *filename)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x4A4, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, filename);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetButtons (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x22D, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetCalculate (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x117, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetChartObjects (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x424, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetCheckBoxes (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x338, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetCheckSpelling (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT customDictionary,
                                              VARIANT ignoreUppercase,
                                              VARIANT alwaysSuggest,
                                              VARIANT spellLang)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x1F9, CAVT_EMPTY,
	                              NULL, 4, __paramTypes, customDictionary,
	                              ignoreUppercase, alwaysSuggest, spellLang);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetClearArrows (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x3CA, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetDrawings (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x304, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetDrawingObjects (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT index,
                                               CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x58, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetDropDowns (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x344, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetEvaluate (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT name,
                                         VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x1, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, name);

	return __result;
}

HRESULT CVIFUNC Excel_Worksheet_Evaluate (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT name,
                                          VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0xFFFFFFFB,
	                              CAVT_VARIANT, returnValue, 1, __paramTypes,
	                              name);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetResetAllPageBreaks (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x592, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetGroupBoxes (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x342, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetGroupObjects (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x459, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetLabels (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x349, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetLines (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x2FF, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetListBoxes (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x340, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetOLEObjects (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x31F, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetOptionButtons (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x33A, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetOvals (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x321, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetPaste (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT destination,
                                      VARIANT link)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0xD3, CAVT_EMPTY, NULL,
	                              2, __paramTypes, destination, link);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetPasteSpecial (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT format, VARIANT link,
                                             VARIANT displayAsIcon,
                                             VARIANT iconFileName,
                                             VARIANT iconIndex,
                                             VARIANT iconLabel)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x403, CAVT_EMPTY,
	                              NULL, 6, __paramTypes, format, link,
	                              displayAsIcon, iconFileName, iconIndex,
	                              iconLabel);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetPictures (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x303, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetPivotTables (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x2B2, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

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
                                                 ExcelObj_PivotTable *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x2AC, CAVT_OBJHANDLE,
	                              returnValue, 16, __paramTypes, sourceType,
	                              sourceData, tableDestination, tableName,
	                              rowGrand, columnGrand, saveData,
	                              hasAutoFormat, autoPage, reserved,
	                              backgroundQuery, optimizeCache,
	                              pageFieldOrder, pageFieldWrapCount, readData,
	                              connection);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetRange (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT cell1,
                                      VARIANT cell2, ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0xC5,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, cell1, cell2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetRectangles (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x306, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetScenarios (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x38C, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetScrollBars (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x33E, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetShowAllData (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x31A, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetShowDataForm (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x199, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetSpinners (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x346, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetTextBoxes (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x309, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetClearCircles (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x59C, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetCircleInvalid (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x59D, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetPrintOut (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT from,
                                         VARIANT to, VARIANT copies,
                                         VARIANT preview, VARIANT activePrinter,
                                         VARIANT printToFile, VARIANT collate,
                                         VARIANT prToFileName)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x6EC, CAVT_EMPTY,
	                              NULL, 8, __paramTypes, from, to, copies,
	                              preview, activePrinter, printToFile, collate,
	                              prToFileName);

	return __result;
}

HRESULT CVIFUNC Excel_Worksheet_CheckSpelling (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT customDictionary,
                                               VARIANT ignoreUppercase,
                                               VARIANT alwaysSuggest,
                                               VARIANT spellLang,
                                               VARIANT ignoreFinalYaa,
                                               VARIANT spellScript)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheet, 0x719, CAVT_EMPTY,
	                              NULL, 6, __paramTypes, customDictionary,
	                              ignoreUppercase, alwaysSuggest, spellLang,
	                              ignoreFinalYaa, spellScript);

	return __result;
}

HRESULT CVIFUNC Excel_NewWorkbook (const char *server, int supportMultithreading,
                                   LCID locale, int reserved,
                                   CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20819, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_CreateObjectByClassIdEx (&clsid, server, &Excel_IID_Workbook,
	                                       supportMultithreading, locale,
	                                       reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_OpenWorkbook (const char *fileName, const char *server,
                                    int supportMultithreading, LCID locale,
                                    int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20819, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_LoadObjectFromFileByClassIdEx (fileName, &clsid, server,
	                                             &Excel_IID_Workbook,
	                                             supportMultithreading, locale,
	                                             reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_ActiveWorkbook (const char *server,
                                      int supportMultithreading, LCID locale,
                                      int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20819, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_GetActiveObjectByClassIdEx (&clsid, server,
	                                          &Excel_IID_Workbook,
	                                          supportMultithreading, locale,
	                                          reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookActivate (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x130, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookChangeFileAccess (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_XlFileAccess mode,
                                                VARIANT writePassword,
                                                VARIANT notify)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x3DD, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              mode, writePassword, notify);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookChangeLink (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, const char *name,
                                          const char *newName,
                                          enum ExcelEnum_XlLinkType type)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x322, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              name, newName, type);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookClose (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT saveChanges,
                                     VARIANT filename, VARIANT routeWorkbook)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x115, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              saveChanges, filename, routeWorkbook);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookGetColors (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x11E, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookSetColors (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         VARIANT newValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x11E, DISPATCH_PROPERTYPUT, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, index, newValue);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookDeleteNumberFormat (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  const char *numberFormat)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x18D, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              numberFormat);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookExclusiveAccess (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VBOOL *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x490, CAVT_BOOL, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookForwardMailer (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x3CD, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookLinkInfo (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, const char *name,
                                        enum ExcelEnum_XlLinkInfo linkInfo,
                                        VARIANT type, VARIANT editionRef,
                                        VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x327, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, name, linkInfo, type,
	                              editionRef);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookLinkSources (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT type,
                                           VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x328, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, type);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookMergeWorkbook (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT filename)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5A6, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              filename);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookNewWindow (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         ExcelObj_Window *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x118, CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookOpenLinks (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, const char *name,
                                         VARIANT readOnly, VARIANT type)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x323, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              name, readOnly, type);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookPivotCaches (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           ExcelObj_PivotCaches *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5A9, CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookPost (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT destName)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x48E, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              destName);

	return __result;
}

HRESULT CVIFUNC Excel_Workbook_PrintOut (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT from,
                                         VARIANT to, VARIANT copies,
                                         VARIANT preview, VARIANT activePrinter,
                                         VARIANT printToFile, VARIANT collate)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x389, CAVT_EMPTY, NULL, 7, __paramTypes,
	                              from, to, copies, preview, activePrinter,
	                              printToFile, collate);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookPrintPreview (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT enableChanges)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x119, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              enableChanges);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookProtect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT password,
                                       VARIANT structure, VARIANT windows)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x11A, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              password, structure, windows);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookProtectSharing (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT filename, VARIANT password,
                                              VARIANT writeResPassword,
                                              VARIANT readOnlyRecommended,
                                              VARIANT createBackup,
                                              VARIANT sharingPassword)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5AA, CAVT_EMPTY, NULL, 6, __paramTypes,
	                              filename, password, writeResPassword,
	                              readOnlyRecommended, createBackup,
	                              sharingPassword);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookRefreshAll (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5AC, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookReply (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x3D1, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookReplyAll (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x3D2, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookRemoveUser (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long index)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5AD, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookRoute (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x3B2, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookRunAutoMacros (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlRunAutoMacro which)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x27A, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              which);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookSave (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x11B, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookSaveAs (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT filename,
                                      VARIANT fileFormat, VARIANT password,
                                      VARIANT writeResPassword,
                                      VARIANT readOnlyRecommended,
                                      VARIANT createBackup,
                                      enum ExcelEnum_XlSaveAsAccessMode accessMode,
                                      VARIANT conflictResolution,
                                      VARIANT addToMru, VARIANT textCodepage,
                                      VARIANT textVisualLayout)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x11C, CAVT_EMPTY, NULL, 11, __paramTypes,
	                              filename, fileFormat, password,
	                              writeResPassword, readOnlyRecommended,
	                              createBackup, accessMode, conflictResolution,
	                              addToMru, textCodepage, textVisualLayout);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookSaveCopyAs (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT filename)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0xAF, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              filename);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookSendMail (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT recipients,
                                        VARIANT subject, VARIANT returnReceipt)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x3B3, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              recipients, subject, returnReceipt);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookSendMailer (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT fileFormat,
                                          enum ExcelEnum_XlPriority priority)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x3D4, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              fileFormat, priority);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookSetLinkOnData (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *name, VARIANT procedure)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x329, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              name, procedure);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookUnprotect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT password)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x11D, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              password);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookUnprotectSharing (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT sharingPassword)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5AF, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              sharingPassword);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookUpdateFromFile (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x3E3, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookUpdateLink (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT name,
                                          VARIANT type)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x324, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              name, type);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookHighlightChangesOptions (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo,
                                                       VARIANT when, VARIANT who,
                                                       VARIANT where)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5B2, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              when, who, where);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookPurgeChangeHistoryNow (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     long days,
                                                     VARIANT sharingPassword)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5B8, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              days, sharingPassword);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookAcceptAllChanges (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT when, VARIANT who,
                                                VARIANT where)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5BA, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              when, who, where);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookRejectAllChanges (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT when, VARIANT who,
                                                VARIANT where)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5BB, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              when, who, where);

	return __result;
}

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
                                                VARIANT connection)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x2AC, CAVT_EMPTY, NULL, 16, __paramTypes,
	                              sourceType, sourceData, tableDestination,
	                              tableName, rowGrand, columnGrand, saveData,
	                              hasAutoFormat, autoPage, reserved,
	                              backgroundQuery, optimizeCache,
	                              pageFieldOrder, pageFieldWrapCount, readData,
	                              connection);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookResetColors (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5BC, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookFollowHyperlink (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               const char *address,
                                               VARIANT subAddress,
                                               VARIANT newWindow,
                                               VARIANT addHistory,
                                               VARIANT extraInfo, VARIANT method,
                                               VARIANT headerInfo)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5BE, CAVT_EMPTY, NULL, 7, __paramTypes,
	                              address, subAddress, newWindow, addHistory,
	                              extraInfo, method, headerInfo);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookAddToFavorites (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x5C4, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookPrintOut (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT from,
                                        VARIANT to, VARIANT copies,
                                        VARIANT preview, VARIANT activePrinter,
                                        VARIANT printToFile, VARIANT collate,
                                        VARIANT prToFileName)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x6EC, CAVT_EMPTY, NULL, 8, __paramTypes,
	                              from, to, copies, preview, activePrinter,
	                              printToFile, collate, prToFileName);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookWebPagePreview (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x71A, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookReloadAs (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_MsoEncoding encoding)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x71D, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              encoding);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbookDummy1 (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, long unused)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x6F6, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              unused);

	return __result;
}

HRESULT CVIFUNC Excel_Workbooksblt (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, const char *s)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Workbook,
	                              0x722, CAVT_EMPTY, NULL, 1, __paramTypes, s);

	return __result;
}

HRESULT CVIFUNC Excel_NewOLEObject (const char *server,
                                    int supportMultithreading, LCID locale,
                                    int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20818, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_CreateObjectByClassIdEx (&clsid, server,
	                                       &Excel_IID_OLEObject,
	                                       supportMultithreading, locale,
	                                       reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_OpenOLEObject (const char *fileName, const char *server,
                                     int supportMultithreading, LCID locale,
                                     int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20818, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_LoadObjectFromFileByClassIdEx (fileName, &clsid, server,
	                                             &Excel_IID_OLEObject,
	                                             supportMultithreading, locale,
	                                             reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_ActiveOLEObject (const char *server,
                                       int supportMultithreading, LCID locale,
                                       int reserved, CAObjHandle *objectHandle)
{
	HRESULT __result;
	GUID clsid = {0x20818, 0x0, 0x0, 0xC0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x0, 0x46};

	__result = CA_GetActiveObjectByClassIdEx (&clsid, server,
	                                          &Excel_IID_OLEObject,
	                                          supportMultithreading, locale,
	                                          reserved, objectHandle);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectBringToFront (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x8001025A,
	                              CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectCopy (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x80010227,
	                              CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectCopyPicture (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlPictureAppearance appearance,
                                            enum ExcelEnum_XlCopyPictureFormat format,
                                            VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x800100D5,
	                              CAVT_VARIANT, returnValue, 2, __paramTypes,
	                              appearance, format);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectCut (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x80010235,
	                              CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x80010075,
	                              CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectDuplicate (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          CAObjHandle *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x8001040F,
	                              CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT replace,
                                       VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x800100EB,
	                              CAVT_VARIANT, returnValue, 1, __paramTypes,
	                              replace);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectSendToBack (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x8001025D,
	                              CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectActivate (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x80010130,
	                              CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectUpdate (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x800102A8,
	                              CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectVerb (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     enum ExcelEnum_XlOLEVerb verb,
                                     VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObject, 0x8001025E,
	                              CAVT_VARIANT, returnValue, 1, __paramTypes,
	                              verb);

	return __result;
}

HRESULT CVIFUNC Excel_AdjustmentsGetItem (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long index,
                                          float *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Adjustments, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_FLOAT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AdjustmentsSetItem (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long index,
                                          float newValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_FLOAT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Adjustments, 0x0,
	                              DISPATCH_PROPERTYPUT, CAVT_EMPTY, NULL, 2,
	                              __paramTypes, index, newValue);

	return __result;
}

HRESULT CVIFUNC Excel_CalloutFormatAutomaticLength (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalloutFormat, 0xA, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CalloutFormatCustomDrop (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, float drop)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalloutFormat, 0xB, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, drop);

	return __result;
}

HRESULT CVIFUNC Excel_CalloutFormatCustomLength (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 float length)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalloutFormat, 0xC, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, length);

	return __result;
}

HRESULT CVIFUNC Excel_CalloutFormatPresetDrop (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               enum ExcelEnum_MsoCalloutDropType dropType)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalloutFormat, 0xD, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, dropType);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeNodesItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_ShapeNode *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeNodes, 0x0, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeNodesDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, long index)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeNodes, 0xB, CAVT_EMPTY, NULL,
	                              1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeNodesInsert (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, long index,
                                        enum ExcelEnum_MsoSegmentType segmentType,
                                        enum ExcelEnum_MsoEditingType editingType,
                                        float x1, float y1, float x2, float y2,
                                        float x3, float y3)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_LONG, CAVT_FLOAT,
	                               CAVT_FLOAT, CAVT_FLOAT, CAVT_FLOAT,
	                               CAVT_FLOAT, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeNodes, 0xC, CAVT_EMPTY, NULL,
	                              9, __paramTypes, index, segmentType,
	                              editingType, x1, y1, x2, y2, x3, y3);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeNodesSetEditingType (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo, long index,
                                                enum ExcelEnum_MsoEditingType editingType)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeNodes, 0xD, CAVT_EMPTY, NULL,
	                              2, __paramTypes, index, editingType);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeNodesSetPosition (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, long index,
                                             float x1, float y1)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_FLOAT, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeNodes, 0xE, CAVT_EMPTY, NULL,
	                              3, __paramTypes, index, x1, y1);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeNodesSetSegmentType (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo, long index,
                                                enum ExcelEnum_MsoSegmentType segmentType)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeNodes, 0xF, CAVT_EMPTY, NULL,
	                              2, __paramTypes, index, segmentType);

	return __result;
}

HRESULT CVIFUNC Excel_PictureFmtIncrementBrightness (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PictureFmt, 0xA, CAVT_EMPTY, NULL,
	                              1, __paramTypes, increment);

	return __result;
}

HRESULT CVIFUNC Excel_PictureFmtIncrementContrast (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PictureFmt, 0xB, CAVT_EMPTY, NULL,
	                              1, __paramTypes, increment);

	return __result;
}

HRESULT CVIFUNC Excel_ShadowFormatIncrementOffsetX (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShadowFormat, 0xA, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, increment);

	return __result;
}

HRESULT CVIFUNC Excel_ShadowFormatIncrementOffsetY (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShadowFormat, 0xB, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, increment);

	return __result;
}

HRESULT CVIFUNC Excel_TextEffectFmtToggleVerticalText (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_TextEffectFmt, 0xA, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ThreeDFmtIncrementRotationX (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ThreeDFmt, 0xA, CAVT_EMPTY, NULL,
	                              1, __paramTypes, increment);

	return __result;
}

HRESULT CVIFUNC Excel_ThreeDFmtIncrementRotationY (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ThreeDFmt, 0xB, CAVT_EMPTY, NULL,
	                              1, __paramTypes, increment);

	return __result;
}

HRESULT CVIFUNC Excel_ThreeDFmtResetRotation (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ThreeDFmt, 0xC, CAVT_EMPTY, NULL,
	                              0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ThreeDFmtSetThreeDFormat (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_MsoPresetThreeDFormat presetThreeDFormat)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ThreeDFmt, 0xD, CAVT_EMPTY, NULL,
	                              1, __paramTypes, presetThreeDFormat);

	return __result;
}

HRESULT CVIFUNC Excel_ThreeDFmtSetExtrusionDirection (CAObjHandle objectHandle,
                                                      ERRORINFO *errorInfo,
                                                      enum ExcelEnum_MsoPresetExtrusionDirection presetExtrusionDirection)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ThreeDFmt, 0xE, CAVT_EMPTY, NULL,
	                              1, __paramTypes, presetExtrusionDirection);

	return __result;
}

HRESULT CVIFUNC Excel_FillFormatBackground (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FillFormat, 0xA, CAVT_EMPTY, NULL,
	                              0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_FillFormatOneColorGradient (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  enum ExcelEnum_MsoGradientStyle style,
                                                  long variant, float degree)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FillFormat, 0xB, CAVT_EMPTY, NULL,
	                              3, __paramTypes, style, variant, degree);

	return __result;
}

HRESULT CVIFUNC Excel_FillFormatPatterned (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_MsoPatternType pattern)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FillFormat, 0xC, CAVT_EMPTY, NULL,
	                              1, __paramTypes, pattern);

	return __result;
}

HRESULT CVIFUNC Excel_FillFormatPresetGradient (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_MsoGradientStyle style,
                                                long variant,
                                                enum ExcelEnum_MsoPresetGradientType presetGradientType)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FillFormat, 0xD, CAVT_EMPTY, NULL,
	                              3, __paramTypes, style, variant,
	                              presetGradientType);

	return __result;
}

HRESULT CVIFUNC Excel_FillFormatPresetTextured (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_MsoPresetTexture presetTexture)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FillFormat, 0xE, CAVT_EMPTY, NULL,
	                              1, __paramTypes, presetTexture);

	return __result;
}

HRESULT CVIFUNC Excel_FillFormatSolid (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FillFormat, 0xF, CAVT_EMPTY, NULL,
	                              0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_FillFormatTwoColorGradient (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  enum ExcelEnum_MsoGradientStyle style,
                                                  long variant)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FillFormat, 0x10, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, style, variant);

	return __result;
}

HRESULT CVIFUNC Excel_FillFormatUserPicture (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *pictureFile)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FillFormat, 0x11, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, pictureFile);

	return __result;
}

HRESULT CVIFUNC Excel_FillFormatUserTextured (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              const char *textureFile)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FillFormat, 0x12, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, textureFile);

	return __result;
}

HRESULT CVIFUNC Excel_SheetsAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT before, VARIANT after, VARIANT count,
                                 VARIANT type, CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0xB5, CAVT_OBJHANDLE, returnValue, 4,
	                              __paramTypes, before, after, count, type);

	return __result;
}

HRESULT CVIFUNC Excel_SheetsCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT before, VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0x227, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              before, after);

	return __result;
}

HRESULT CVIFUNC Excel_SheetsDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0x75, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SheetsFillAcrossSheets (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              ExcelObj_Range range,
                                              enum ExcelEnum_XlFillWith type)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0x1D5, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              range, type);

	return __result;
}

HRESULT CVIFUNC Excel_SheetsItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_SheetsMove (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT before, VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0x27D, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              before, after);

	return __result;
}

HRESULT CVIFUNC Excel_Sheets_PrintOut (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT from,
                                       VARIANT to, VARIANT copies,
                                       VARIANT preview, VARIANT activePrinter,
                                       VARIANT printToFile, VARIANT collate)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0x389, CAVT_EMPTY, NULL, 7, __paramTypes,
	                              from, to, copies, preview, activePrinter,
	                              printToFile, collate);

	return __result;
}

HRESULT CVIFUNC Excel_SheetsPrintPreview (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT enableChanges)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0x119, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              enableChanges);

	return __result;
}

HRESULT CVIFUNC Excel_SheetsSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT replace)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0xEB, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              replace);

	return __result;
}

HRESULT CVIFUNC Excel_Sheets_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_SheetsPrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate,
                                      VARIANT prToFileName)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Sheets,
	                              0x6EC, CAVT_EMPTY, NULL, 8, __paramTypes,
	                              from, to, copies, preview, activePrinter,
	                              printToFile, collate, prToFileName);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbooksAdd (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT template,
                                    ExcelObj_Workbook *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Workbooks, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, template);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbooksClose (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Workbooks, 0x115, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbooksItem (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     ExcelObj_Workbook *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Workbooks, 0xAA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbooksOpen (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, const char *filename,
                                     VARIANT updateLinks, VARIANT readOnly,
                                     VARIANT format, VARIANT password,
                                     VARIANT writeResPassword,
                                     VARIANT ignoreReadOnlyRecommended,
                                     VARIANT origin, VARIANT delimiter,
                                     VARIANT editable, VARIANT notify,
                                     VARIANT converter, VARIANT addToMru,
                                     ExcelObj_Workbook *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Workbooks, 0x2AA, CAVT_OBJHANDLE,
	                              returnValue, 13, __paramTypes, filename,
	                              updateLinks, readOnly, format, password,
	                              writeResPassword, ignoreReadOnlyRecommended,
	                              origin, delimiter, editable, notify,
	                              converter, addToMru);

	return __result;
}

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
                                          VARIANT textVisualLayout)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Workbooks, 0x2AB, CAVT_EMPTY,
	                              NULL, 14, __paramTypes, filename, origin,
	                              startRow, dataType, textQualifier,
	                              consecutiveDelimiter, tab, semicolon, comma,
	                              space, other, otherChar, fieldInfo,
	                              textVisualLayout);

	return __result;
}

HRESULT CVIFUNC Excel_Workbooks_Default (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         ExcelObj_Workbook *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Workbooks, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

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
                                         VARIANT thousandsSeparator)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Workbooks, 0x6ED, CAVT_EMPTY,
	                              NULL, 16, __paramTypes, filename, origin,
	                              startRow, dataType, textQualifier,
	                              consecutiveDelimiter, tab, semicolon, comma,
	                              space, other, otherChar, fieldInfo,
	                              textVisualLayout, decimalSeparator,
	                              thousandsSeparator);

	return __result;
}

HRESULT CVIFUNC Excel_PublishObjectDelete (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PublishObject, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PublishObjectPublish (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT create)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PublishObject, 0x767, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, create);

	return __result;
}

HRESULT CVIFUNC Excel_WebOptionsUseDefaultFolderSuffix (CAObjHandle objectHandle,
                                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WebOptions, 0x77B, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CubeFieldsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_CubeField *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_CubeFields, 0xAA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_CubeFields_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_CubeField *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_CubeFields, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WindowActivate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x130, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WindowActivateNext (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x45B, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WindowActivatePrevious (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x45C, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WindowClose (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT saveChanges,
                                   VARIANT filename, VARIANT routeWorkbook,
                                   VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x115, CAVT_BOOL, returnValue, 3,
	                              __paramTypes, saveChanges, filename,
	                              routeWorkbook);

	return __result;
}

HRESULT CVIFUNC Excel_WindowLargeScroll (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT down,
                                         VARIANT up, VARIANT toRight,
                                         VARIANT toLeft, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x223, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, down, up, toRight, toLeft);

	return __result;
}

HRESULT CVIFUNC Excel_WindowNewWindow (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       ExcelObj_Window *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x118, CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WindowPrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate,
                                      VARIANT prToFileName, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x6EC, CAVT_VARIANT, returnValue, 8,
	                              __paramTypes, from, to, copies, preview,
	                              activePrinter, printToFile, collate,
	                              prToFileName);

	return __result;
}

HRESULT CVIFUNC Excel_WindowPrintPreview (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT enableChanges,
                                          VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x119, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, enableChanges);

	return __result;
}

HRESULT CVIFUNC Excel_WindowScrollWorkbookTabs (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT sheets, VARIANT position,
                                                VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x296, CAVT_VARIANT, returnValue, 2,
	                              __paramTypes, sheets, position);

	return __result;
}

HRESULT CVIFUNC Excel_WindowSmallScroll (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT down,
                                         VARIANT up, VARIANT toRight,
                                         VARIANT toLeft, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x224, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, down, up, toRight, toLeft);

	return __result;
}

HRESULT CVIFUNC Excel_WindowPointsToScreenPixelsX (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   long points,
                                                   long *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x6F0, CAVT_LONG, returnValue, 1,
	                              __paramTypes, points);

	return __result;
}

HRESULT CVIFUNC Excel_WindowPointsToScreenPixelsY (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   long points,
                                                   long *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x6F1, CAVT_LONG, returnValue, 1,
	                              __paramTypes, points);

	return __result;
}

HRESULT CVIFUNC Excel_WindowRangeFromPoint (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, long x, long y,
                                            CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x6F2, CAVT_OBJHANDLE, returnValue, 2,
	                              __paramTypes, x, y);

	return __result;
}

HRESULT CVIFUNC Excel_WindowScrollIntoView (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, long left,
                                            long top, long width, long height,
                                            VARIANT start)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_LONG, CAVT_LONG,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Window,
	                              0x6F5, CAVT_EMPTY, NULL, 5, __paramTypes,
	                              left, top, width, height, start);

	return __result;
}

HRESULT CVIFUNC Excel_WindowsArrange (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_XlArrangeStyle arrangeStyle,
                                      VARIANT activeWorkbook,
                                      VARIANT syncHorizontal,
                                      VARIANT syncVertical, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Windows,
	                              0x27E, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, arrangeStyle, activeWorkbook,
	                              syncHorizontal, syncVertical);

	return __result;
}

HRESULT CVIFUNC Excel_WindowsItem (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT index,
                                   ExcelObj_Window *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Windows,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Windows_Default (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_Window *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Windows,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

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
                                                    VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0xA9,
	                              CAVT_VARIANT, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                              VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4000,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIsNA (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4002,
	                              CAVT_BOOL, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIsError (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4003,
	                              CAVT_BOOL, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

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
                                            VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4004,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4005,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                            VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4006,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                            VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4007,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                            VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x400B,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                              VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x400C,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDollar (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               VARIANT arg2, char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x400D,
	                              CAVT_CSTRING, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFixed (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x400E,
	                              CAVT_CSTRING, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPi (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           double *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4013,
	                              CAVT_DOUBLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionLn (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, double arg1,
                                           double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4016,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionLog10 (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4017,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionRound (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x401B,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionLookup (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3,
                                               VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x401C,
	                              CAVT_VARIANT, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIndex (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              double arg2, VARIANT arg3,
                                              VARIANT arg4, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_DOUBLE, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x401D,
	                              CAVT_VARIANT, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionRept (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *arg1, double arg2,
                                             char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x401E,
	                              CAVT_CSTRING, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

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
                                            VARIANT arg30, VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4024,
	                              CAVT_BOOL, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                           VARIANT arg30, VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4025,
	                              CAVT_BOOL, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDCount (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               ExcelObj_Range arg1, VARIANT arg2,
                                               VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4028,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDSum (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4029,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDAverage (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Range arg1,
                                                 VARIANT arg2, VARIANT arg3,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x402A,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDMin (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x402B,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDMax (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x402C,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDStDev (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               ExcelObj_Range arg1, VARIANT arg2,
                                               VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x402D,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

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
                                            VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x402E,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDVar (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x402F,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionText (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             const char *arg2,
                                             char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4030,
	                              CAVT_CSTRING, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionLinEst (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4031,
	                              CAVT_VARIANT, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionTrend (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              VARIANT arg4, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4032,
	                              CAVT_VARIANT, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionLogEst (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4033,
	                              CAVT_VARIANT, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionGrowth (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4034,
	                              CAVT_VARIANT, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPv (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, double arg1,
                                           double arg2, double arg3,
                                           VARIANT arg4, VARIANT arg5,
                                           double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4038,
	                              CAVT_DOUBLE, returnValue, 5, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFv (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, double arg1,
                                           double arg2, double arg3,
                                           VARIANT arg4, VARIANT arg5,
                                           double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4039,
	                              CAVT_DOUBLE, returnValue, 5, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionNPer (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             VARIANT arg4, VARIANT arg5,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x403A,
	                              CAVT_DOUBLE, returnValue, 5, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPmt (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            VARIANT arg4, VARIANT arg5,
                                            double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x403B,
	                              CAVT_DOUBLE, returnValue, 5, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionRate (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             VARIANT arg4, VARIANT arg5,
                                             VARIANT arg6, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x403C,
	                              CAVT_DOUBLE, returnValue, 6, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionMIrr (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             double arg2, double arg3,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x403D,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIrr (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT arg1,
                                            VARIANT arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x403E,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionMatch (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4040,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionWeekday (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4046,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSearch (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               const char *arg1,
                                               const char *arg2, VARIANT arg3,
                                               double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4052,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionTranspose (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1,
                                                  VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4053,
	                              CAVT_VARIANT, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionAtan2 (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4061,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionAsin (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4062,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionAcos (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4063,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

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
                                               VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4064,
	                              CAVT_VARIANT, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionHLookup (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3, VARIANT arg4,
                                                VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4065,
	                              CAVT_VARIANT, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionVLookup (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3, VARIANT arg4,
                                                VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4066,
	                              CAVT_VARIANT, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionLog (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            VARIANT arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x406D,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionProper (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               const char *arg1,
                                               char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4072,
	                              CAVT_CSTRING, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionTrim (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *arg1,
                                             char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4076,
	                              CAVT_CSTRING, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionReplace (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                const char *arg1, double arg2,
                                                double arg3, const char *arg4,
                                                char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4077,
	                              CAVT_CSTRING, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSubstitute (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   const char *arg1,
                                                   const char *arg2,
                                                   const char *arg3,
                                                   VARIANT arg4,
                                                   char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING, CAVT_CSTRING,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4078,
	                              CAVT_CSTRING, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFind (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *arg1, const char *arg2,
                                             VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x407C,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIsErr (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x407E,
	                              CAVT_BOOL, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIsText (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x407F,
	                              CAVT_BOOL, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIsNumber (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1,
                                                 VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4080,
	                              CAVT_BOOL, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSln (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x408E,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSyd (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            double arg4, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x408F,
	                              CAVT_DOUBLE, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDdb (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            double arg4, VARIANT arg5,
                                            double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_DOUBLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4090,
	                              CAVT_DOUBLE, returnValue, 5, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionClean (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              const char *arg1,
                                              char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40A2,
	                              CAVT_CSTRING, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionMDeterm (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40A3,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionMInverse (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1,
                                                 VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40A4,
	                              CAVT_VARIANT, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionMMult (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40A5,
	                              CAVT_VARIANT, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIpmt (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             double arg4, VARIANT arg5,
                                             VARIANT arg6, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_DOUBLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40A7,
	                              CAVT_DOUBLE, returnValue, 6, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPpmt (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             double arg4, VARIANT arg5,
                                             VARIANT arg6, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_DOUBLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40A8,
	                              CAVT_DOUBLE, returnValue, 6, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6);

	return __result;
}

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
                                               double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40A9,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40B7,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFact (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40B8,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDProduct (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Range arg1,
                                                 VARIANT arg2, VARIANT arg3,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40BD,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIsNonText (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1,
                                                  VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40BE,
	                              CAVT_BOOL, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

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
                                               double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40C1,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                             VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40C2,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDStDevP (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                ExcelObj_Range arg1,
                                                VARIANT arg2, VARIANT arg3,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40C3,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDVarP (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              ExcelObj_Range arg1, VARIANT arg2,
                                              VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40C4,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIsLogical (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1,
                                                  VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40C6,
	                              CAVT_BOOL, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDCountA (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                ExcelObj_Range arg1,
                                                VARIANT arg2, VARIANT arg3,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40C7,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionUSDollar (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, double arg2,
                                                 char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40CC,
	                              CAVT_CSTRING, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFindB (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              const char *arg1, const char *arg2,
                                              VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40CD,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSearchB (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                const char *arg1,
                                                const char *arg2, VARIANT arg3,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40CE,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionReplaceB (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 const char *arg1, double arg2,
                                                 double arg3, const char *arg4,
                                                 char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40CF,
	                              CAVT_CSTRING, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionRoundUp (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40D4,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionRoundDown (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40D5,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionRank (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             ExcelObj_Range arg2, VARIANT arg3,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_OBJHANDLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40D8,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDays360 (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40DC,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionVdb (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            double arg4, double arg5,
                                            VARIANT arg6, VARIANT arg7,
                                            double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_DOUBLE, CAVT_DOUBLE, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40DE,
	                              CAVT_DOUBLE, returnValue, 7, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7);

	return __result;
}

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
                                               double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40E3,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                                   double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40E4,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSinh (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40E5,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionCosh (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40E6,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionTanh (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40E7,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionAsinh (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40E8,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionAcosh (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40E9,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionAtanh (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40EA,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDGet (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40EB,
	                              CAVT_VARIANT, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDb (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, double arg1,
                                           double arg2, double arg3, double arg4,
                                           VARIANT arg5, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_DOUBLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40F7,
	                              CAVT_DOUBLE, returnValue, 5, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFrequency (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1, VARIANT arg2,
                                                  VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40FC,
	                              CAVT_VARIANT, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

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
                                               double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x410D,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionBetaDist (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, double arg2,
                                                 double arg3, VARIANT arg4,
                                                 VARIANT arg5,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x410E,
	                              CAVT_DOUBLE, returnValue, 5, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionGammaLn (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x410F,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionBetaInv (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double arg3, VARIANT arg4,
                                                VARIANT arg5,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4110,
	                              CAVT_DOUBLE, returnValue, 5, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionBinomDist (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  double arg3, VBOOL arg4,
                                                  double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_BOOL};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4111,
	                              CAVT_DOUBLE, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionChiDist (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4112,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionChiInv (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4113,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionCombin (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4114,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionConfidence (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   double arg1, double arg2,
                                                   double arg3,
                                                   double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4115,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionCritBinom (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  double arg3,
                                                  double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4116,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionEven (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4117,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionExponDist (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  VBOOL arg3,
                                                  double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_BOOL};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4118,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFDist (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double arg3,
                                              double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4119,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFInv (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x411A,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFisher (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x411B,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFisherInv (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1,
                                                  double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x411C,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFloor (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x411D,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionGammaDist (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  double arg3, VBOOL arg4,
                                                  double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_BOOL};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x411E,
	                              CAVT_DOUBLE, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionGammaInv (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, double arg2,
                                                 double arg3,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x411F,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionCeiling (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4120,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionHypGeomDist (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    double arg1, double arg2,
                                                    double arg3, double arg4,
                                                    double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4121,
	                              CAVT_DOUBLE, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionLogNormDist (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    double arg1, double arg2,
                                                    double arg3,
                                                    double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4122,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionLogInv (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double arg2, double arg3,
                                               double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4123,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionNegBinomDist (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     double arg1, double arg2,
                                                     double arg3,
                                                     double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4124,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionNormDist (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, double arg2,
                                                 double arg3, VBOOL arg4,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_BOOL};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4125,
	                              CAVT_DOUBLE, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionNormSDist (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1,
                                                  double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4126,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionNormInv (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4127,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionNormSInv (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4128,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionStandardize (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    double arg1, double arg2,
                                                    double arg3,
                                                    double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4129,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionOdd (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x412A,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPermut (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x412B,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPoisson (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                VBOOL arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_BOOL};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x412C,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionTDist (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double arg3,
                                              double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x412D,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionWeibull (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double arg3, VBOOL arg4,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_BOOL};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x412E,
	                              CAVT_DOUBLE, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSumXMY2 (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x412F,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSumX2MY2 (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1, VARIANT arg2,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4130,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSumX2PY2 (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1, VARIANT arg2,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4131,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionChiTest (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4132,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionCorrel (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4133,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionCovar (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4134,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionForecast (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, VARIANT arg2,
                                                 VARIANT arg3,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4135,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionFTest (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4136,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIntercept (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1, VARIANT arg2,
                                                  double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4137,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPearson (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4138,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionRSq (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT arg1,
                                            VARIANT arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4139,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionStEyx (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x413A,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSlope (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x413B,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionTTest (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double arg3,
                                              double arg4, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_DOUBLE,
	                               CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x413C,
	                              CAVT_DOUBLE, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionProb (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             VARIANT arg2, double arg3,
                                             VARIANT arg4, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_DOUBLE,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x413D,
	                              CAVT_DOUBLE, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

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
                                              VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x413E,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x413F,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4140,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                              VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4141,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                             VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4142,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

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
                                             VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4143,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionZTest (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              double arg2, VARIANT arg3,
                                              double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_DOUBLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4144,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionLarge (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4145,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSmall (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4146,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionQuartile (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1, double arg2,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4147,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPercentile (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   VARIANT arg1, double arg2,
                                                   double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4148,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPercentRank (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    VARIANT arg1, double arg2,
                                                    VARIANT arg3,
                                                    double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_DOUBLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4149,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

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
                                             VARIANT arg30, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x414A,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionTrimMean (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1, double arg2,
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x414B,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionTInv (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x414C,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPower (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4151,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionRadians (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4156,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDegrees (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4157,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

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
                                                 double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_OBJHANDLE, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4158,
	                              CAVT_DOUBLE, returnValue, 30, __paramTypes,
	                              arg1, arg2, arg3, arg4, arg5, arg6, arg7,
	                              arg8, arg9, arg10, arg11, arg12, arg13,
	                              arg14, arg15, arg16, arg17, arg18, arg19,
	                              arg20, arg21, arg22, arg23, arg24, arg25,
	                              arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionSumIf (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              ExcelObj_Range arg1, VARIANT arg2,
                                              VARIANT arg3, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4159,
	                              CAVT_DOUBLE, returnValue, 3, __paramTypes,
	                              arg1, arg2, arg3);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionCountIf (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                ExcelObj_Range arg1,
                                                VARIANT arg2,
                                                double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x415A,
	                              CAVT_DOUBLE, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionCountBlank (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   ExcelObj_Range arg1,
                                                   double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x415B,
	                              CAVT_DOUBLE, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionIspmt (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double arg3,
                                              double arg4, double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x415E,
	                              CAVT_DOUBLE, returnValue, 4, __paramTypes,
	                              arg1, arg2, arg3, arg4);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionRoman (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              VARIANT arg2, char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4162,
	                              CAVT_CSTRING, returnValue, 2, __paramTypes,
	                              arg1, arg2);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionAsc (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            const char *arg1, char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40D6,
	                              CAVT_CSTRING, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionDbcs (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *arg1,
                                             char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x40D7,
	                              CAVT_CSTRING, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetFunctionPhonetic (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Range arg1,
                                                 char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_WorksheetFunction, 0x4168,
	                              CAVT_CSTRING, returnValue, 1, __paramTypes,
	                              arg1);

	return __result;
}

HRESULT CVIFUNC Excel_RangeActivate (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x130, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAddress (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT rowAbsolute,
                                    VARIANT columnAbsolute,
                                    enum ExcelEnum_XlReferenceStyle referenceStyle,
                                    VARIANT external, VARIANT relativeTo,
                                    char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xEC, DISPATCH_PROPERTYGET, CAVT_CSTRING,
	                              returnValue, 5, __paramTypes, rowAbsolute,
	                              columnAbsolute, referenceStyle, external,
	                              relativeTo);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAddressLocal (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT rowAbsolute,
                                         VARIANT columnAbsolute,
                                         enum ExcelEnum_XlReferenceStyle referenceStyle,
                                         VARIANT external, VARIANT relativeTo,
                                         char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1B5, DISPATCH_PROPERTYGET, CAVT_CSTRING,
	                              returnValue, 5, __paramTypes, rowAbsolute,
	                              columnAbsolute, referenceStyle, external,
	                              relativeTo);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAdvancedFilter (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_XlFilterAction action,
                                           VARIANT criteriaRange,
                                           VARIANT copyToRange, VARIANT unique,
                                           VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x36C, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, action, criteriaRange,
	                              copyToRange, unique);

	return __result;
}

HRESULT CVIFUNC Excel_RangeApplyNames (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT names,
                                       VARIANT ignoreRelativeAbsolute,
                                       VARIANT useRowColumnNames,
                                       VARIANT omitColumn, VARIANT omitRow,
                                       enum ExcelEnum_XlApplyNamesOrder order,
                                       VARIANT appendLast, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_LONG,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1B9, CAVT_VARIANT, returnValue, 7,
	                              __paramTypes, names, ignoreRelativeAbsolute,
	                              useRowColumnNames, omitColumn, omitRow,
	                              order, appendLast);

	return __result;
}

HRESULT CVIFUNC Excel_RangeApplyOutlineStyles (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1C0, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAutoComplete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         const char *string, char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x4A1, CAVT_CSTRING, returnValue, 1,
	                              __paramTypes, string);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAutoFill (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     ExcelObj_Range destination,
                                     enum ExcelEnum_XlAutoFillType type,
                                     VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1C1, CAVT_VARIANT, returnValue, 2,
	                              __paramTypes, destination, type);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAutoFilter (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT field,
                                       VARIANT criteria1,
                                       enum ExcelEnum_XlAutoFilterOperator operator,
                                       VARIANT criteria2,
                                       VARIANT visibleDropDown,
                                       VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x319, CAVT_VARIANT, returnValue, 5,
	                              __paramTypes, field, criteria1, operator,
	                              criteria2, visibleDropDown);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAutoFit (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xED, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAutoFormat (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       enum ExcelEnum_XlRangeAutoFormat format,
                                       VARIANT number, VARIANT font,
                                       VARIANT alignment, VARIANT border,
                                       VARIANT pattern, VARIANT width,
                                       VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x72, CAVT_VARIANT, returnValue, 7,
	                              __paramTypes, format, number, font,
	                              alignment, border, pattern, width);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAutoOutline (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x40C, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeBorderAround (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT lineStyle,
                                         enum ExcelEnum_XlBorderWeight weight,
                                         enum ExcelEnum_XlColorIndex colorIndex,
                                         VARIANT color, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG, CAVT_LONG,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x42B, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, lineStyle, weight, colorIndex,
	                              color);

	return __result;
}

HRESULT CVIFUNC Excel_RangeCalculate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x117, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeCharacters (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT start,
                                       VARIANT length,
                                       ExcelObj_Characters *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x25B, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, start, length);

	return __result;
}

HRESULT CVIFUNC Excel_RangeCheckSpelling (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT customDictionary,
                                          VARIANT ignoreUppercase,
                                          VARIANT alwaysSuggest,
                                          VARIANT spellLang,
                                          VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1F9, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, customDictionary,
	                              ignoreUppercase, alwaysSuggest, spellLang);

	return __result;
}

HRESULT CVIFUNC Excel_RangeClear (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x6F, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeClearContents (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x71, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeClearFormats (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x70, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeClearNotes (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xEF, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeClearOutline (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x40D, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeColumnDifferences (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT comparison,
                                              ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1FE, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, comparison);

	return __result;
}

HRESULT CVIFUNC Excel_RangeConsolidate (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT sources,
                                        VARIANT function, VARIANT topRow,
                                        VARIANT leftColumn, VARIANT createLinks,
                                        VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1E2, CAVT_VARIANT, returnValue, 5,
	                              __paramTypes, sources, function, topRow,
	                              leftColumn, createLinks);

	return __result;
}

HRESULT CVIFUNC Excel_RangeCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT destination, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x227, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, destination);

	return __result;
}

HRESULT CVIFUNC Excel_RangeCopyFromRecordset (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              LPUNKNOWN data, VARIANT maxRows,
                                              VARIANT maxColumns,
                                              long *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_UNKNOWN, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x480, CAVT_LONG, returnValue, 3,
	                              __paramTypes, data, maxRows, maxColumns);

	return __result;
}

HRESULT CVIFUNC Excel_RangeCopyPicture (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_XlPictureAppearance appearance,
                                        enum ExcelEnum_XlCopyPictureFormat format,
                                        VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xD5, CAVT_VARIANT, returnValue, 2,
	                              __paramTypes, appearance, format);

	return __result;
}

HRESULT CVIFUNC Excel_RangeCreateNames (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT top,
                                        VARIANT left, VARIANT bottom,
                                        VARIANT right, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1C9, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, top, left, bottom, right);

	return __result;
}

HRESULT CVIFUNC Excel_RangeCreatePublisher (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT edition,
                                            enum ExcelEnum_XlPictureAppearance appearance,
                                            VARIANT containsPICT,
                                            VARIANT containsBIFF,
                                            VARIANT containsRTF,
                                            VARIANT containsVALU,
                                            VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1CA, CAVT_VARIANT, returnValue, 6,
	                              __paramTypes, edition, appearance,
	                              containsPICT, containsBIFF, containsRTF,
	                              containsVALU);

	return __result;
}

HRESULT CVIFUNC Excel_RangeCut (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT destination, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x235, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, destination);

	return __result;
}

HRESULT CVIFUNC Excel_RangeDataSeries (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT rowcol,
                                       enum ExcelEnum_XlDataSeriesType type,
                                       enum ExcelEnum_XlDataSeriesDate date,
                                       VARIANT step, VARIANT stop, VARIANT trend,
                                       VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1D0, CAVT_VARIANT, returnValue, 6,
	                              __paramTypes, rowcol, type, date, step, stop,
	                              trend);

	return __result;
}

HRESULT CVIFUNC Excel_RangeGet_Default (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT rowIndex,
                                        VARIANT columnIndex,
                                        VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, rowIndex,
	                              columnIndex);

	return __result;
}

HRESULT CVIFUNC Excel_RangeSet_Default (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT rowIndex,
                                        VARIANT columnIndex, VARIANT newValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x0, DISPATCH_PROPERTYPUT, CAVT_EMPTY, NULL,
	                              3, __paramTypes, rowIndex, columnIndex,
	                              newValue);

	return __result;
}

HRESULT CVIFUNC Excel_RangeDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT shift,
                                   VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x75, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, shift);

	return __result;
}

HRESULT CVIFUNC Excel_RangeDialogBox (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xF5, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeEditionOptions (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_XlEditionType type,
                                           enum ExcelEnum_XlEditionOptionsOption option,
                                           VARIANT name, VARIANT reference,
                                           enum ExcelEnum_XlPictureAppearance appearance,
                                           enum ExcelEnum_XlPictureAppearance chartSize,
                                           VARIANT format, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_LONG, CAVT_LONG,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x46B, CAVT_VARIANT, returnValue, 7,
	                              __paramTypes, type, option, name, reference,
	                              appearance, chartSize, format);

	return __result;
}

HRESULT CVIFUNC Excel_RangeEnd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                enum ExcelEnum_XlDirection direction,
                                ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1F4, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, direction);

	return __result;
}

HRESULT CVIFUNC Excel_RangeFillDown (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xF8, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeFillLeft (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xF9, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeFillRight (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xFA, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeFillUp (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xFB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeFind (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT what, VARIANT after, VARIANT lookIn,
                                 VARIANT lookAt, VARIANT searchOrder,
                                 enum ExcelEnum_XlSearchDirection searchDirection,
                                 VARIANT matchCase, VARIANT matchByte,
                                 ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x18E, CAVT_OBJHANDLE, returnValue, 8,
	                              __paramTypes, what, after, lookIn, lookAt,
	                              searchOrder, searchDirection, matchCase,
	                              matchByte);

	return __result;
}

HRESULT CVIFUNC Excel_RangeFindNext (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT after,
                                     ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x18F, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, after);

	return __result;
}

HRESULT CVIFUNC Excel_RangeFindPrevious (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT after,
                                         ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x190, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, after);

	return __result;
}

HRESULT CVIFUNC Excel_RangeFunctionWizard (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x23B, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeGoalSeek (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT goal,
                                     ExcelObj_Range changingCell,
                                     VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_OBJHANDLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1D8, CAVT_BOOL, returnValue, 2,
	                              __paramTypes, goal, changingCell);

	return __result;
}

HRESULT CVIFUNC Excel_RangeGroup (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT start, VARIANT end, VARIANT by,
                                  VARIANT periods, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x2E, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, start, end, by, periods);

	return __result;
}

HRESULT CVIFUNC Excel_RangeInsertIndent (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, long insertAmount)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x565, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              insertAmount);

	return __result;
}

HRESULT CVIFUNC Excel_RangeInsert (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT shift,
                                   VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xFC, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, shift);

	return __result;
}

HRESULT CVIFUNC Excel_RangeGetItem (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT rowIndex,
                                    VARIANT columnIndex, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, rowIndex,
	                              columnIndex);

	return __result;
}

HRESULT CVIFUNC Excel_RangeSetItem (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT rowIndex,
                                    VARIANT columnIndex, VARIANT newValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xAA, DISPATCH_PROPERTYPUT, CAVT_EMPTY, NULL,
	                              3, __paramTypes, rowIndex, columnIndex,
	                              newValue);

	return __result;
}

HRESULT CVIFUNC Excel_RangeJustify (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1EF, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeListNames (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xFD, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeMerge (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT across)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x234, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              across);

	return __result;
}

HRESULT CVIFUNC Excel_RangeUnMerge (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x568, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeNavigateArrow (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT towardPrecedent,
                                          VARIANT arrowNumber,
                                          VARIANT linkNumber,
                                          VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x408, CAVT_VARIANT, returnValue, 3,
	                              __paramTypes, towardPrecedent, arrowNumber,
	                              linkNumber);

	return __result;
}

HRESULT CVIFUNC Excel_RangeNoteText (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT text,
                                     VARIANT start, VARIANT length,
                                     char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x467, CAVT_CSTRING, returnValue, 3,
	                              __paramTypes, text, start, length);

	return __result;
}

HRESULT CVIFUNC Excel_RangeOffset (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT rowOffset,
                                   VARIANT columnOffset,
                                   ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xFE, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, rowOffset,
	                              columnOffset);

	return __result;
}

HRESULT CVIFUNC Excel_RangeParse (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT parseLine, VARIANT destination,
                                  VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1DD, CAVT_VARIANT, returnValue, 2,
	                              __paramTypes, parseLine, destination);

	return __result;
}

HRESULT CVIFUNC Excel_RangePasteSpecial (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlPasteType paste,
                                         enum ExcelEnum_XlPasteSpecialOperation operation,
                                         VARIANT skipBlanks, VARIANT transpose,
                                         VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x403, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, paste, operation, skipBlanks,
	                              transpose);

	return __result;
}

HRESULT CVIFUNC Excel_Range_PrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate,
                                      VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x389, CAVT_VARIANT, returnValue, 7,
	                              __paramTypes, from, to, copies, preview,
	                              activePrinter, printToFile, collate);

	return __result;
}

HRESULT CVIFUNC Excel_RangePrintPreview (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT enableChanges,
                                         VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x119, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, enableChanges);

	return __result;
}

HRESULT CVIFUNC Excel_RangeRange (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT cell1, VARIANT cell2,
                                  ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xC5, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, cell1, cell2);

	return __result;
}

HRESULT CVIFUNC Excel_RangeRemoveSubtotal (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x373, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeReplace (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT what,
                                    VARIANT replacement, VARIANT lookAt,
                                    VARIANT searchOrder, VARIANT matchCase,
                                    VARIANT matchByte, VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xE2, CAVT_BOOL, returnValue, 6,
	                              __paramTypes, what, replacement, lookAt,
	                              searchOrder, matchCase, matchByte);

	return __result;
}

HRESULT CVIFUNC Excel_RangeResize (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT rowSize,
                                   VARIANT columnSize,
                                   ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x100, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, rowSize,
	                              columnSize);

	return __result;
}

HRESULT CVIFUNC Excel_RangeRowDifferences (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT comparison,
                                           ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1FF, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, comparison);

	return __result;
}

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
                                VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x103, CAVT_VARIANT, returnValue, 30,
	                              __paramTypes, arg1, arg2, arg3, arg4, arg5,
	                              arg6, arg7, arg8, arg9, arg10, arg11, arg12,
	                              arg13, arg14, arg15, arg16, arg17, arg18,
	                              arg19, arg20, arg21, arg22, arg23, arg24,
	                              arg25, arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_RangeSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeShow (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1F0, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeShowDependents (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT remove,
                                           VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x36D, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, remove);

	return __result;
}

HRESULT CVIFUNC Excel_RangeShowErrors (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x36E, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeShowPrecedents (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT remove,
                                           VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x36F, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, remove);

	return __result;
}

HRESULT CVIFUNC Excel_RangeSort (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT key1, enum ExcelEnum_XlSortOrder order1,
                                 VARIANT key2, VARIANT type,
                                 enum ExcelEnum_XlSortOrder order2, VARIANT key3,
                                 enum ExcelEnum_XlSortOrder order3,
                                 enum ExcelEnum_XlYesNoGuess header,
                                 VARIANT orderCustom, VARIANT matchCase,
                                 enum ExcelEnum_XlSortOrientation orientation,
                                 enum ExcelEnum_XlSortMethod sortMethod,
                                 VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_LONG, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x370, CAVT_VARIANT, returnValue, 12,
	                              __paramTypes, key1, order1, key2, type,
	                              order2, key3, order3, header, orderCustom,
	                              matchCase, orientation, sortMethod);

	return __result;
}

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
                                        VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_LONG, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x371, CAVT_VARIANT, returnValue, 12,
	                              __paramTypes, sortMethod, key1, order1, type,
	                              key2, order2, key3, order3, header,
	                              orderCustom, matchCase, orientation);

	return __result;
}

HRESULT CVIFUNC Excel_RangeSpecialCells (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlCellType type,
                                         VARIANT value,
                                         ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x19A, CAVT_OBJHANDLE, returnValue, 2,
	                              __paramTypes, type, value);

	return __result;
}

HRESULT CVIFUNC Excel_RangeSubscribeTo (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        const char *edition,
                                        enum ExcelEnum_XlSubscribeToFormat format,
                                        VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1E1, CAVT_VARIANT, returnValue, 2,
	                              __paramTypes, edition, format);

	return __result;
}

HRESULT CVIFUNC Excel_RangeSubtotal (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long groupBy,
                                     enum ExcelEnum_XlConsolidationFunction function,
                                     VARIANT totalList, VARIANT replace,
                                     VARIANT pageBreaks,
                                     enum ExcelEnum_XlSummaryRow summaryBelowData,
                                     VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x372, CAVT_VARIANT, returnValue, 6,
	                              __paramTypes, groupBy, function, totalList,
	                              replace, pageBreaks, summaryBelowData);

	return __result;
}

HRESULT CVIFUNC Excel_RangeTable (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT rowInput, VARIANT columnInput,
                                  VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x1F1, CAVT_VARIANT, returnValue, 2,
	                              __paramTypes, rowInput, columnInput);

	return __result;
}

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
                                          VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x410, CAVT_VARIANT, returnValue, 13,
	                              __paramTypes, destination, dataType,
	                              textQualifier, consecutiveDelimiter, tab,
	                              semicolon, comma, space, other, otherChar,
	                              fieldInfo, decimalSeparator,
	                              thousandsSeparator);

	return __result;
}

HRESULT CVIFUNC Excel_RangeUngroup (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0xF4, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeAddComment (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT text,
                                       ExcelObj_Comment *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x56D, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, text);

	return __result;
}

HRESULT CVIFUNC Excel_RangeClearComments (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x56E, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangeSetPhonetic (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x714, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RangePrintOut (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT from,
                                     VARIANT to, VARIANT copies, VARIANT preview,
                                     VARIANT activePrinter, VARIANT printToFile,
                                     VARIANT collate, VARIANT prToFileName,
                                     VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Range,
	                              0x6EC, CAVT_VARIANT, returnValue, 8,
	                              __paramTypes, from, to, copies, preview,
	                              activePrinter, printToFile, collate,
	                              prToFileName);

	return __result;
}

HRESULT CVIFUNC Excel_VPageBreakDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_VPageBreak, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_VPageBreakDragOff (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlDirection direction,
                                         long regionIndex)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_VPageBreak, 0x58C, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, direction,
	                              regionIndex);

	return __result;
}

HRESULT CVIFUNC Excel_HPageBreakDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_HPageBreak, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_HPageBreakDragOff (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlDirection direction,
                                         long regionIndex)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_HPageBreak, 0x58C, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, direction,
	                              regionIndex);

	return __result;
}

HRESULT CVIFUNC Excel_HPageBreaksItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_HPageBreak *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_HPageBreaks, 0xAA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_HPageBreaks_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long index,
                                           ExcelObj_HPageBreak *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_HPageBreaks, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_HPageBreaksAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, CAObjHandle before,
                                      ExcelObj_HPageBreak *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_HPageBreaks, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, before);

	return __result;
}

HRESULT CVIFUNC Excel_VPageBreaksItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_VPageBreak *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_VPageBreaks, 0xAA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_VPageBreaks_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long index,
                                           ExcelObj_VPageBreak *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_VPageBreaks, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_VPageBreaksAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, CAObjHandle before,
                                      ExcelObj_VPageBreak *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_VPageBreaks, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, before);

	return __result;
}

HRESULT CVIFUNC Excel_RecentFileOpen (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      ExcelObj_Workbook *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_RecentFile, 0x2AA, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RecentFileDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_RecentFile, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RecentFilesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_RecentFile *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_RecentFiles, 0xAA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_RecentFiles_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long index,
                                           ExcelObj_RecentFile *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_RecentFiles, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_RecentFilesAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, const char *name,
                                      ExcelObj_RecentFile *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_RecentFiles, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, name);

	return __result;
}

HRESULT CVIFUNC Excel_StyleDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Style,
	                              0x75, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_StylesAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 const char *name, VARIANT basedOn,
                                 ExcelObj_Style *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Styles,
	                              0xB5, CAVT_OBJHANDLE, returnValue, 2,
	                              __paramTypes, name, basedOn);

	return __result;
}

HRESULT CVIFUNC Excel_StylesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, ExcelObj_Style *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Styles,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_StylesMerge (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT workbook,
                                   VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Styles,
	                              0x234, CAVT_VARIANT, returnValue, 1,
	                              __paramTypes, workbook);

	return __result;
}

HRESULT CVIFUNC Excel_Styles_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Style *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Styles,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_BordersItem (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo,
                                   enum ExcelEnum_XlBordersIndex index,
                                   ExcelObj_Border *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Borders,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Borders_Default (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       enum ExcelEnum_XlBordersIndex index,
                                       ExcelObj_Border *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Borders,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AddInsAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 const char *filename, VARIANT copyFile,
                                 ExcelObj_AddIn *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_AddIns,
	                              0xB5, CAVT_OBJHANDLE, returnValue, 2,
	                              __paramTypes, filename, copyFile);

	return __result;
}

HRESULT CVIFUNC Excel_AddInsItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, ExcelObj_AddIn *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_AddIns,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AddIns_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_AddIn *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_AddIns,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AreasItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 long index, ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Areas,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Areas_Default (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long index,
                                     ExcelObj_Range *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Areas,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartsAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT before, VARIANT after, VARIANT count,
                                 ExcelObj_Chart *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0xB5, CAVT_OBJHANDLE, returnValue, 3,
	                              __paramTypes, before, after, count);

	return __result;
}

HRESULT CVIFUNC Excel_ChartsCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT before, VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0x227, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              before, after);

	return __result;
}

HRESULT CVIFUNC Excel_ChartsDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0x75, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_Charts_Dummy7 (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0x10007, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartsItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartsMove (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT before, VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0x27D, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              before, after);

	return __result;
}

HRESULT CVIFUNC Excel_Charts_PrintOut (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT from,
                                       VARIANT to, VARIANT copies,
                                       VARIANT preview, VARIANT activePrinter,
                                       VARIANT printToFile, VARIANT collate)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0x389, CAVT_EMPTY, NULL, 7, __paramTypes,
	                              from, to, copies, preview, activePrinter,
	                              printToFile, collate);

	return __result;
}

HRESULT CVIFUNC Excel_ChartsPrintPreview (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT enableChanges)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0x119, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              enableChanges);

	return __result;
}

HRESULT CVIFUNC Excel_ChartsSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT replace)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0xEB, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              replace);

	return __result;
}

HRESULT CVIFUNC Excel_Charts_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartsPrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate,
                                      VARIANT prToFileName)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Charts,
	                              0x6EC, CAVT_EMPTY, NULL, 8, __paramTypes,
	                              from, to, copies, preview, activePrinter,
	                              printToFile, collate, prToFileName);

	return __result;
}

HRESULT CVIFUNC Excel_PivotCacheRefresh (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotCache, 0x589, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotCacheResetTimer (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotCache, 0x72A, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotCacheCreatePivotTable (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT tableDestination,
                                                  VARIANT tableName,
                                                  VARIANT readData,
                                                  ExcelObj_PivotTable *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotCache, 0x72C, CAVT_OBJHANDLE,
	                              returnValue, 3, __paramTypes,
	                              tableDestination, tableName, readData);

	return __result;
}

HRESULT CVIFUNC Excel_PivotCachesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_PivotCache *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotCaches, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotCaches_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           ExcelObj_PivotCache *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotCaches, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotCaches_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotCaches, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotCachesAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_XlPivotTableSourceType sourceType,
                                      VARIANT sourceData,
                                      ExcelObj_PivotCache *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotCaches, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, sourceType,
	                              sourceData);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFormulaDelete (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotFormula, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFormulasAdd (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        const char *formula,
                                        ExcelObj_PivotFormula *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotFormulas, 0xB5,
	                              CAVT_OBJHANDLE, returnValue, 1, __paramTypes,
	                              formula);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFormulasItem (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         ExcelObj_PivotFormula *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotFormulas, 0xAA,
	                              CAVT_OBJHANDLE, returnValue, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFormulas_Default (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             ExcelObj_PivotFormula *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotFormulas, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFormulas_NewEnum (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotFormulas, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableAddFields (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT rowFields,
                                           VARIANT columnFields,
                                           VARIANT pageFields,
                                           VARIANT addToTable,
                                           VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2C4, CAVT_VARIANT,
	                              returnValue, 4, __paramTypes, rowFields,
	                              columnFields, pageFields, addToTable);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableColumnFields (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2C9,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableShowPages (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT pageField,
                                           VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2C2, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, pageField);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableDataFields (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2CB,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableHiddenFields (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2C7,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTablePageFields (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2CA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTablePivotFields (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2CE, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableRefreshTable (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VBOOL *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2CD, CAVT_BOOL,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableRowFields (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2C8,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableVisibleFields (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT index,
                                               CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2C6,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableCalculatedFields (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  ExcelObj_CalculatedFields *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x5CC, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableGetData (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, const char *name,
                                         double *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x5D3, CAVT_DOUBLE,
	                              returnValue, 1, __paramTypes, name);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableListFormulas (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x5D4, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTablePivotCache (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            ExcelObj_PivotCache *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x5D8, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

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
                                                  VARIANT connection)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2AC, CAVT_EMPTY,
	                              NULL, 16, __paramTypes, sourceType,
	                              sourceData, tableDestination, tableName,
	                              rowGrand, columnGrand, saveData,
	                              hasAutoFormat, autoPage, reserved,
	                              backgroundQuery, optimizeCache,
	                              pageFieldOrder, pageFieldWrapCount, readData,
	                              connection);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTablePivotSelect (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *name,
                                             enum ExcelEnum_XlPTSelectionMode mode)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x5DD, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, name, mode);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableUpdate (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x2A8, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTableFormat (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_xlPivotFormatType format)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTable, 0x74, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, format);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTablesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_PivotTable *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTables, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTables_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTables, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotTablesAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      ExcelObj_PivotCache pivotCache,
                                      VARIANT tableDestination,
                                      VARIANT tableName, VARIANT readData,
                                      ExcelObj_PivotTable *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotTables, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 4, __paramTypes, pivotCache,
	                              tableDestination, tableName, readData);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldChildItems (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x2DA,
	                              DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldHiddenItems (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x2D8,
	                              DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldParentItems (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x2D9,
	                              DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldPivotItems (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x2E1, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldGetSubtotals (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x2DD,
	                              DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldSetSubtotals (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index, VARIANT newValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x2DD,
	                              DISPATCH_PROPERTYPUT, CAVT_EMPTY, NULL, 2,
	                              __paramTypes, index, newValue);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldVisibleItems (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x2D7,
	                              DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldCalculatedItems (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_CalculatedItems *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x5E3, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldAutoSort (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long order,
                                          const char *field)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x5EA, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, order, field);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldAutoShow (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long type,
                                          long range, long count,
                                          const char *field)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_LONG,
	                               CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotField, 0x5EB, CAVT_EMPTY,
	                              NULL, 4, __paramTypes, type, range, count,
	                              field);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFieldsItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotFields, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotFields_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotFields, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CalculatedFieldsAdd (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           const char *name, const char *formula,
                                           ExcelObj_PivotField *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalculatedFields, 0xB5,
	                              CAVT_OBJHANDLE, returnValue, 2, __paramTypes,
	                              name, formula);

	return __result;
}

HRESULT CVIFUNC Excel_CalculatedFieldsItem (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            ExcelObj_PivotField *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalculatedFields, 0xAA,
	                              CAVT_OBJHANDLE, returnValue, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_CalculatedFields_Default (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT field,
                                                ExcelObj_PivotField *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_CalculatedFields, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, field);

	return __result;
}

HRESULT CVIFUNC Excel_CalculatedFields_NewEnum (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalculatedFields, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotItemChildItems (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotItem, 0x2DA,
	                              DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotItemDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotItem, 0x75, CAVT_EMPTY, NULL,
	                              0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PivotItemsAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, const char *name)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotItems, 0xB5, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, name);

	return __result;
}

HRESULT CVIFUNC Excel_PivotItemsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotItems, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotItems_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotItems, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CalculatedItemsAdd (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, const char *name,
                                          const char *formula,
                                          ExcelObj_PivotItem *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalculatedItems, 0xB5,
	                              CAVT_OBJHANDLE, returnValue, 2, __paramTypes,
	                              name, formula);

	return __result;
}

HRESULT CVIFUNC Excel_CalculatedItemsItem (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           ExcelObj_PivotItem *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalculatedItems, 0xAA,
	                              CAVT_OBJHANDLE, returnValue, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_CalculatedItems_Default (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT field,
                                               ExcelObj_PivotItem *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_CalculatedItems, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, field);

	return __result;
}

HRESULT CVIFUNC Excel_CalculatedItems_NewEnum (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CalculatedItems, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CharactersDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Characters, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CharactersInsert (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, const char *string,
                                        VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Characters, 0xFC, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, string);

	return __result;
}

HRESULT CVIFUNC Excel_DialogsItem (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo,
                                   enum ExcelEnum_XlBuiltInDialog index,
                                   ExcelObj_Dialog *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Dialogs,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Dialogs_Default (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       enum ExcelEnum_XlBuiltInDialog index,
                                       ExcelObj_Dialog *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Dialogs,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

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
                                  VBOOL *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Dialog,
	                              0x1F0, CAVT_BOOL, returnValue, 30,
	                              __paramTypes, arg1, arg2, arg3, arg4, arg5,
	                              arg6, arg7, arg8, arg9, arg10, arg11, arg12,
	                              arg13, arg14, arg15, arg16, arg17, arg18,
	                              arg19, arg20, arg21, arg22, arg23, arg24,
	                              arg25, arg26, arg27, arg28, arg29, arg30);

	return __result;
}

HRESULT CVIFUNC Excel_SoundNoteDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SoundNote, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SoundNoteImport (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       const char *filename,
                                       VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SoundNote, 0x395, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, filename);

	return __result;
}

HRESULT CVIFUNC Excel_SoundNotePlay (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SoundNote, 0x396, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SoundNoteRecord (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SoundNote, 0x397, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PanesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 long index, ExcelObj_Pane *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Panes,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Panes_Default (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long index,
                                     ExcelObj_Pane *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Panes,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PaneActivate (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VBOOL *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Pane,
	                              0x130, CAVT_BOOL, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PaneLargeScroll (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT down,
                                       VARIANT up, VARIANT toRight,
                                       VARIANT toLeft, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Pane,
	                              0x223, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, down, up, toRight, toLeft);

	return __result;
}

HRESULT CVIFUNC Excel_PaneSmallScroll (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT down,
                                       VARIANT up, VARIANT toRight,
                                       VARIANT toLeft, VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Pane,
	                              0x224, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, down, up, toRight, toLeft);

	return __result;
}

HRESULT CVIFUNC Excel_PaneScrollIntoView (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long left,
                                          long top, long width, long height,
                                          VARIANT start)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_LONG, CAVT_LONG,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Pane,
	                              0x6F5, CAVT_EMPTY, NULL, 5, __paramTypes,
	                              left, top, width, height, start);

	return __result;
}

HRESULT CVIFUNC Excel_ScenariosAdd (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, const char *name,
                                    VARIANT changingCells, VARIANT values,
                                    VARIANT comment, VARIANT locked,
                                    VARIANT hidden,
                                    ExcelObj_Scenario *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Scenarios, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 6, __paramTypes, name,
	                              changingCells, values, comment, locked,
	                              hidden);

	return __result;
}

HRESULT CVIFUNC Excel_ScenariosCreateSummary (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              enum ExcelEnum_XlSummaryReportType reportType,
                                              VARIANT resultCells,
                                              VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Scenarios, 0x391, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, reportType,
	                              resultCells);

	return __result;
}

HRESULT CVIFUNC Excel_ScenariosItem (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     ExcelObj_Scenario *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Scenarios, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ScenariosMerge (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT source,
                                      VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Scenarios, 0x234, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, source);

	return __result;
}

HRESULT CVIFUNC Excel_Scenarios_NewEnum (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Scenarios, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ScenarioChangeScenario (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT changingCells,
                                              VARIANT values,
                                              VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Scenario,
	                              0x390, CAVT_VARIANT, returnValue, 2,
	                              __paramTypes, changingCells, values);

	return __result;
}

HRESULT CVIFUNC Excel_ScenarioDelete (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Scenario,
	                              0x75, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ScenarioShow (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Scenario,
	                              0x1F0, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ScenarioValues (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Scenario,
	                              0xA4, DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy3 (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10003, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsBringToFront (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x25A, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsCopy (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x227, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsCopyPicture (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlPictureAppearance appearance,
                                             enum ExcelEnum_XlCopyPictureFormat format,
                                             VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0xD5, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, appearance,
	                              format);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsCut (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x235, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsDuplicate (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x40F, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy12 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x1000C, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy15 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x1000F, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT replace,
                                        VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0xEB, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, replace);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsSendToBack (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x25D, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy22 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10016, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy30 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x1001E, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy32 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10020, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy33 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10021, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy34 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10022, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy36 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10024, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy37 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10025, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy38 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10026, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy39 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10027, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy40 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10028, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_Dummy41 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x10029, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT classType,
                                     VARIANT filename, VARIANT link,
                                     VARIANT displayAsIcon, VARIANT iconFileName,
                                     VARIANT iconIndex, VARIANT iconLabel,
                                     VARIANT left, VARIANT top, VARIANT width,
                                     VARIANT height,
                                     ExcelObj_OLEObject *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 11, __paramTypes, classType,
	                              filename, link, displayAsIcon, iconFileName,
	                              iconIndex, iconLabel, left, top, width,
	                              height);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsGroup (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       ExcelObj_GroupObject *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0x2E, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjects_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEObjects, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_RoutingSlipGetRecipients (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index,
                                                VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_RoutingSlip, 0x3B8,
	                              DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_RoutingSlipSetRecipients (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index, VARIANT newValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_RoutingSlip, 0x3B8,
	                              DISPATCH_PROPERTYPUT, CAVT_EMPTY, NULL, 2,
	                              __paramTypes, index, newValue);

	return __result;
}

HRESULT CVIFUNC Excel_RoutingSlipReset (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_RoutingSlip, 0x22B, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OutlineShowLevels (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT rowLevels,
                                         VARIANT columnLevels,
                                         VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Outline,
	                              0x3C0, CAVT_VARIANT, returnValue, 2,
	                              __paramTypes, rowLevels, columnLevels);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetsAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT before,
                                     VARIANT after, VARIANT count, VARIANT type,
                                     CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 4, __paramTypes, before, after,
	                              count, type);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetsCopy (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT before,
                                      VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0x227, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, before, after);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetsDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetsFillAcrossSheets (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  ExcelObj_Range range,
                                                  enum ExcelEnum_XlFillWith type)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0x1D5, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, range, type);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0xAA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetsMove (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT before,
                                      VARIANT after)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0x27D, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, before, after);

	return __result;
}

HRESULT CVIFUNC Excel_Worksheets_PrintOut (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT from,
                                           VARIANT to, VARIANT copies,
                                           VARIANT preview,
                                           VARIANT activePrinter,
                                           VARIANT printToFile, VARIANT collate)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0x389, CAVT_EMPTY,
	                              NULL, 7, __paramTypes, from, to, copies,
	                              preview, activePrinter, printToFile, collate);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetsPrintPreview (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT enableChanges)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0x119, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, enableChanges);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetsSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT replace)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0xEB, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, replace);

	return __result;
}

HRESULT CVIFUNC Excel_Worksheets_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_WorksheetsPrintOut (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT from,
                                          VARIANT to, VARIANT copies,
                                          VARIANT preview, VARIANT activePrinter,
                                          VARIANT printToFile, VARIANT collate,
                                          VARIANT prToFileName)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Worksheets, 0x6EC, CAVT_EMPTY,
	                              NULL, 8, __paramTypes, from, to, copies,
	                              preview, activePrinter, printToFile, collate,
	                              prToFileName);

	return __result;
}

HRESULT CVIFUNC Excel_PageSetupGetPrintQuality (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index,
                                                VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PageSetup, 0x3FE,
	                              DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PageSetupSetPrintQuality (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index, VARIANT newValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PageSetup, 0x3FE,
	                              DISPATCH_PROPERTYPUT, CAVT_EMPTY, NULL, 2,
	                              __paramTypes, index, newValue);

	return __result;
}

HRESULT CVIFUNC Excel_NamesAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT name, VARIANT refersTo, VARIANT visible,
                                VARIANT macroType, VARIANT shortcutKey,
                                VARIANT category, VARIANT nameLocal,
                                VARIANT refersToLocal, VARIANT categoryLocal,
                                VARIANT refersToR1C1, VARIANT refersToR1C1Local,
                                ExcelObj_Name *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Names,
	                              0xB5, CAVT_OBJHANDLE, returnValue, 11,
	                              __paramTypes, name, refersTo, visible,
	                              macroType, shortcutKey, category, nameLocal,
	                              refersToLocal, categoryLocal, refersToR1C1,
	                              refersToR1C1Local);

	return __result;
}

HRESULT CVIFUNC Excel_NamesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT index, VARIANT indexLocal,
                                 VARIANT refersTo, ExcelObj_Name *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Names,
	                              0xAA, CAVT_OBJHANDLE, returnValue, 3,
	                              __paramTypes, index, indexLocal, refersTo);

	return __result;
}

HRESULT CVIFUNC Excel_Names_Default (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     VARIANT indexLocal, VARIANT refersTo,
                                     ExcelObj_Name *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Names,
	                              0x0, CAVT_OBJHANDLE, returnValue, 3,
	                              __paramTypes, index, indexLocal, refersTo);

	return __result;
}

HRESULT CVIFUNC Excel_NameDelete (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Name,
	                              0x75, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectBringToFront (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObject, 0x25A, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectCopy (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObject, 0x227, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectCopyPicture (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              enum ExcelEnum_XlPictureAppearance appearance,
                                              enum ExcelEnum_XlCopyPictureFormat format,
                                              VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObject, 0xD5, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, appearance,
	                              format);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectCut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObject, 0x235, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectDelete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObject, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectDuplicate (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            CAObjHandle *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObject, 0x40F,
	                              CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectSelect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT replace,
                                         VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObject, 0xEB, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, replace);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectSendToBack (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObject, 0x25D, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectActivate (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObject, 0x130, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjects_Dummy3 (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x10003, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsBringToFront (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x25A, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsCopy (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x227, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsCopyPicture (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               enum ExcelEnum_XlPictureAppearance appearance,
                                               enum ExcelEnum_XlCopyPictureFormat format,
                                               VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0xD5, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, appearance,
	                              format);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsCut (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x235, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsDelete (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsDuplicate (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             CAObjHandle *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x40F,
	                              CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjects_Dummy12 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x1000C, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjects_Dummy15 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x1000F, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsSelect (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT replace,
                                          VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0xEB, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, replace);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsSendToBack (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x25D, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjects_Dummy22 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x10016, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjects_Dummy25 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x10019, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjects_Dummy27 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x1001B, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjects_Dummy28 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x1001C, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjects_Dummy29 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x1001D, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsAdd (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, double left,
                                       double top, double width, double height,
                                       ExcelObj_ChartObject *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_DOUBLE, CAVT_DOUBLE, CAVT_DOUBLE,
	                               CAVT_DOUBLE};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0xB5,
	                              CAVT_OBJHANDLE, returnValue, 4, __paramTypes,
	                              left, top, width, height);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsGroup (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         ExcelObj_GroupObject *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0x2E,
	                              CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjectsItem (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0xAA,
	                              CAVT_OBJHANDLE, returnValue, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartObjects_NewEnum (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartObjects, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CustomViewsItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT viewName,
                                       ExcelObj_CustomView *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CustomViews, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, viewName);

	return __result;
}

HRESULT CVIFUNC Excel_CustomViewsAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, const char *viewName,
                                      VARIANT printSettings,
                                      VARIANT rowColSettings,
                                      ExcelObj_CustomView *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CustomViews, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 3, __paramTypes, viewName,
	                              printSettings, rowColSettings);

	return __result;
}

HRESULT CVIFUNC Excel_CustomViews_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT viewName,
                                           ExcelObj_CustomView *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_CustomViews, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, viewName);

	return __result;
}

HRESULT CVIFUNC Excel_CustomViewShow (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CustomView, 0x1F0, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CustomViewDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_CustomView, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_FormatConditionsItem (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            ExcelObj_FormatCondition *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FormatConditions, 0xAA,
	                              CAVT_OBJHANDLE, returnValue, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_FormatConditionsAdd (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_XlFormatConditionType type,
                                           VARIANT operator, VARIANT formula1,
                                           VARIANT formula2,
                                           ExcelObj_FormatCondition *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FormatConditions, 0xB5,
	                              CAVT_OBJHANDLE, returnValue, 4, __paramTypes,
	                              type, operator, formula1, formula2);

	return __result;
}

HRESULT CVIFUNC Excel_FormatConditions_Default (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index,
                                                ExcelObj_FormatCondition *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_FormatConditions, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_FormatConditionsDelete (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FormatConditions, 0x75,
	                              CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_FormatConditionModify (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlFormatConditionType type,
                                             VARIANT operator, VARIANT formula1,
                                             VARIANT formula2)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FormatCondition, 0x62D,
	                              CAVT_EMPTY, NULL, 4, __paramTypes, type,
	                              operator, formula1, formula2);

	return __result;
}

HRESULT CVIFUNC Excel_FormatConditionDelete (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FormatCondition, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CommentsItem (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, long index,
                                    ExcelObj_Comment *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Comments,
	                              0xAA, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Comments_Default (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, long index,
                                        ExcelObj_Comment *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Comments,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_CommentText (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT text,
                                   VARIANT start, VARIANT overwrite,
                                   char **returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Comment,
	                              0x8A, CAVT_CSTRING, returnValue, 3,
	                              __paramTypes, text, start, overwrite);

	return __result;
}

HRESULT CVIFUNC Excel_CommentDelete (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Comment,
	                              0x75, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CommentNext (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo,
                                   ExcelObj_Comment *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Comment,
	                              0x1F6, CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CommentPrevious (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       ExcelObj_Comment *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Comment,
	                              0x1F7, CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_QueryTablesAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT connection,
                                      ExcelObj_Range destination, VARIANT sql,
                                      ExcelObj_QueryTable *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_OBJHANDLE, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_QueryTables, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 3, __paramTypes, connection,
	                              destination, sql);

	return __result;
}

HRESULT CVIFUNC Excel_QueryTablesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_QueryTable *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_QueryTables, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_QueryTables_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           ExcelObj_QueryTable *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_QueryTables, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_QueryTables_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_QueryTables, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ParameterSetParam (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlParameterType type,
                                         VARIANT value)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Parameter, 0x641, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, type, value);

	return __result;
}

HRESULT CVIFUNC Excel_ParametersAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, const char *name,
                                     VARIANT iDataType,
                                     ExcelObj_Parameter *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Parameters, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, name,
	                              iDataType);

	return __result;
}

HRESULT CVIFUNC Excel_ParametersItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Parameter *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Parameters, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Parameters_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_Parameter *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Parameters, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ParametersDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Parameters, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_Parameters_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Parameters, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ODBCErrorsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, long index,
                                      ExcelObj_ODBCError *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ODBCErrors, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ODBCErrors_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long index,
                                          ExcelObj_ODBCError *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_ODBCErrors, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ODBCErrors_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ODBCErrors, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ValidationAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     enum ExcelEnum_XlDVType type,
                                     VARIANT alertStyle, VARIANT operator,
                                     VARIANT formula1, VARIANT formula2)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Validation, 0xB5, CAVT_EMPTY,
	                              NULL, 5, __paramTypes, type, alertStyle,
	                              operator, formula1, formula2);

	return __result;
}

HRESULT CVIFUNC Excel_ValidationDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Validation, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ValidationModify (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT type,
                                        VARIANT alertStyle, VARIANT operator,
                                        VARIANT formula1, VARIANT formula2)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Validation, 0x62D, CAVT_EMPTY,
	                              NULL, 5, __paramTypes, type, alertStyle,
	                              operator, formula1, formula2);

	return __result;
}

HRESULT CVIFUNC Excel_HyperlinksAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, CAObjHandle anchor,
                                     const char *address, VARIANT subAddress,
                                     VARIANT screenTip, VARIANT textToDisplay,
                                     CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_CSTRING, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Hyperlinks, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 5, __paramTypes, anchor,
	                              address, subAddress, screenTip,
	                              textToDisplay);

	return __result;
}

HRESULT CVIFUNC Excel_HyperlinksItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Hyperlink *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Hyperlinks, 0xAA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Hyperlinks_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_Hyperlink *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Hyperlinks, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_HyperlinksDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Hyperlinks, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_HyperlinkAddToFavorites (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Hyperlink, 0x5C4, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_HyperlinkDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Hyperlink, 0x75, CAVT_EMPTY, NULL,
	                              0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_HyperlinkFollow (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT newWindow,
                                       VARIANT addHistory, VARIANT extraInfo,
                                       VARIANT method, VARIANT headerInfo)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Hyperlink, 0x650, CAVT_EMPTY,
	                              NULL, 5, __paramTypes, newWindow, addHistory,
	                              extraInfo, method, headerInfo);

	return __result;
}

HRESULT CVIFUNC Excel_HyperlinkCreateNewDocument (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  const char *filename,
                                                  VBOOL editNow, VBOOL overwrite)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_BOOL, CAVT_BOOL};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Hyperlink, 0x75C, CAVT_EMPTY,
	                              NULL, 3, __paramTypes, filename, editNow,
	                              overwrite);

	return __result;
}

HRESULT CVIFUNC Excel_Filters_Default (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_Filter *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Filters,
	                              0x0, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_FiltersItem (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, long index,
                                   ExcelObj_Filter *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Filters,
	                              0xAA, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AutoCorrectAddReplacement (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 const char *what,
                                                 const char *replacement,
                                                 VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_AutoCorrect, 0x47A, CAVT_VARIANT,
	                              returnValue, 2, __paramTypes, what,
	                              replacement);

	return __result;
}

HRESULT CVIFUNC Excel_AutoCorrectDeleteReplacement (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    const char *what,
                                                    VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_AutoCorrect, 0x47B, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, what);

	return __result;
}

HRESULT CVIFUNC Excel_AutoCorrectGetReplacementList (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     VARIANT index,
                                                     VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_AutoCorrect, 0x47F,
	                              DISPATCH_PROPERTYGET, CAVT_VARIANT,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_AutoCorrectSetReplacementList (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     VARIANT index,
                                                     VARIANT newValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_AutoCorrect, 0x47F,
	                              DISPATCH_PROPERTYPUT, CAVT_EMPTY, NULL, 2,
	                              __paramTypes, index, newValue);

	return __result;
}

HRESULT CVIFUNC Excel_ChartFillFormatOneColorGradient (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo,
                                                       enum ExcelEnum_MsoGradientStyle style,
                                                       long variant,
                                                       float degree)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartFillFormat, 0x655,
	                              CAVT_EMPTY, NULL, 3, __paramTypes, style,
	                              variant, degree);

	return __result;
}

HRESULT CVIFUNC Excel_ChartFillFormatTwoColorGradient (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo,
                                                       enum ExcelEnum_MsoGradientStyle style,
                                                       long variant)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartFillFormat, 0x658,
	                              CAVT_EMPTY, NULL, 2, __paramTypes, style,
	                              variant);

	return __result;
}

HRESULT CVIFUNC Excel_ChartFillFormatPresetTextured (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     enum ExcelEnum_MsoPresetTexture presetTexture)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartFillFormat, 0x659,
	                              CAVT_EMPTY, NULL, 1, __paramTypes,
	                              presetTexture);

	return __result;
}

HRESULT CVIFUNC Excel_ChartFillFormatSolid (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartFillFormat, 0x65B,
	                              CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartFillFormatPatterned (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_MsoPatternType pattern)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartFillFormat, 0x65C,
	                              CAVT_EMPTY, NULL, 1, __paramTypes, pattern);

	return __result;
}

HRESULT CVIFUNC Excel_ChartFillFormatUserPicture (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT pictureFile,
                                                  VARIANT pictureFormat,
                                                  VARIANT pictureStackUnit,
                                                  VARIANT picturePlacement)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartFillFormat, 0x65D,
	                              CAVT_EMPTY, NULL, 4, __paramTypes,
	                              pictureFile, pictureFormat, pictureStackUnit,
	                              picturePlacement);

	return __result;
}

HRESULT CVIFUNC Excel_ChartFillFormatUserTextured (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   const char *textureFile)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartFillFormat, 0x662,
	                              CAVT_EMPTY, NULL, 1, __paramTypes,
	                              textureFile);

	return __result;
}

HRESULT CVIFUNC Excel_ChartFillFormatPresetGradient (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     enum ExcelEnum_MsoGradientStyle style,
                                                     long variant,
                                                     enum ExcelEnum_MsoPresetGradientType presetGradientType)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartFillFormat, 0x664,
	                              CAVT_EMPTY, NULL, 3, __paramTypes, style,
	                              variant, presetGradientType);

	return __result;
}

HRESULT CVIFUNC Excel_AxisDelete (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Axis,
	                              0x75, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AxisSelect (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Axis,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartTitleSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartTitle, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartTitleDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartTitle, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartTitleCharacters (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT start,
                                            VARIANT length,
                                            ExcelObj_Characters *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartTitle, 0x25B,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, start, length);

	return __result;
}

HRESULT CVIFUNC Excel_AxisTitleSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_AxisTitle, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AxisTitleDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_AxisTitle, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AxisTitleCharacters (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT start,
                                           VARIANT length,
                                           ExcelObj_Characters *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_AxisTitle, 0x25B,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, start, length);

	return __result;
}

HRESULT CVIFUNC Excel_ChartGroupSeriesCollection (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT index,
                                                  CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartGroup, 0x44, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartGroupsItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_ChartGroup *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartGroups, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ChartGroups_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartGroups, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_AxesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                enum ExcelEnum_XlAxisType type,
                                enum ExcelEnum_XlAxisGroup axisGroup,
                                ExcelObj_Axis *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Axes,
	                              0xAA, CAVT_OBJHANDLE, returnValue, 2,
	                              __paramTypes, type, axisGroup);

	return __result;
}

HRESULT CVIFUNC Excel_Axes_NewEnum (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Axes,
	                              0xFFFFFFFC, CAVT_UNKNOWN, returnValue, 0,
	                              NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PointsItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  long index, ExcelObj_Point *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Points,
	                              0xAA, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Points_NewEnum (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Points,
	                              0xFFFFFFFC, CAVT_UNKNOWN, returnValue, 0,
	                              NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PointApplyDataLabels (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlDataLabelsType type,
                                            VARIANT legendKey, VARIANT autoText,
                                            VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Point,
	                              0x97, CAVT_VARIANT, returnValue, 3,
	                              __paramTypes, type, legendKey, autoText);

	return __result;
}

HRESULT CVIFUNC Excel_PointClearFormats (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Point,
	                              0x70, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PointCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Point,
	                              0x227, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PointDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Point,
	                              0x75, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PointPaste (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Point,
	                              0xD3, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PointSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Point,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesApplyDataLabels (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlDataLabelsType type,
                                             VARIANT legendKey, VARIANT autoText,
                                             VARIANT hasLeaderLines,
                                             VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0x97, CAVT_VARIANT, returnValue, 4,
	                              __paramTypes, type, legendKey, autoText,
	                              hasLeaderLines);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesClearFormats (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0x70, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0x227, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesDataLabels (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0x9D, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0x75, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesErrorBar (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_XlErrorBarDirection direction,
                                      enum ExcelEnum_XlErrorBarInclude include,
                                      enum ExcelEnum_XlErrorBarType type,
                                      VARIANT amount, VARIANT minusValues,
                                      VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_LONG,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0x98, CAVT_VARIANT, returnValue, 5,
	                              __paramTypes, direction, include, type,
	                              amount, minusValues);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesPaste (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0xD3, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesPoints (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT index,
                                    CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0x46, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesTrendlines (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0x9A, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesApplyCustomType (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlChartType chartType)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Series,
	                              0x579, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              chartType);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesCollectionAdd (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT source,
                                           enum ExcelEnum_XlRowCol rowcol,
                                           VARIANT seriesLabels,
                                           VARIANT categoryLabels,
                                           VARIANT replace,
                                           ExcelObj_Series *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_LONG, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SeriesCollection, 0xB5,
	                              CAVT_OBJHANDLE, returnValue, 5, __paramTypes,
	                              source, rowcol, seriesLabels, categoryLabels,
	                              replace);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesCollectionExtend (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT source, VARIANT rowcol,
                                              VARIANT categoryLabels,
                                              VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SeriesCollection, 0xE3,
	                              CAVT_VARIANT, returnValue, 3, __paramTypes,
	                              source, rowcol, categoryLabels);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesCollectionItem (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            ExcelObj_Series *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SeriesCollection, 0xAA,
	                              CAVT_OBJHANDLE, returnValue, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesCollection_NewEnum (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SeriesCollection, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesCollectionPaste (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlRowCol rowcol,
                                             VARIANT seriesLabels,
                                             VARIANT categoryLabels,
                                             VARIANT replace, VARIANT newSeries,
                                             VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SeriesCollection, 0xD3,
	                              CAVT_VARIANT, returnValue, 5, __paramTypes,
	                              rowcol, seriesLabels, categoryLabels,
	                              replace, newSeries);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesCollectionNewSeries (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Series *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SeriesCollection, 0x45D,
	                              CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabelSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabel, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabelDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabel, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabelCharacters (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT start,
                                           VARIANT length,
                                           ExcelObj_Characters *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabel, 0x25B,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, start, length);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabelsSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabels, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabelsDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabels, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabels_Dummy9 (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabels, 0x10009, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabels_Dummy10 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabels, 0x1000A, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabels_Dummy13 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabels, 0x1000D, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabels_Dummy16 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabels, 0x10010, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabels_Dummy17 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabels, 0x10011, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabelsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_DataLabel *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabels, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_DataLabels_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataLabels, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LegendEntryDelete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LegendEntry, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LegendEntrySelect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LegendEntry, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LegendEntriesItem (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         ExcelObj_LegendEntry *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LegendEntries, 0xAA,
	                              CAVT_OBJHANDLE, returnValue, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_LegendEntries_NewEnum (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LegendEntries, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LegendKeyClearFormats (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LegendKey, 0x70, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LegendKeyDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LegendKey, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LegendKeySelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LegendKey, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_TrendlinesAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     enum ExcelEnum_XlTrendlineType type,
                                     VARIANT order, VARIANT period,
                                     VARIANT forward, VARIANT backward,
                                     VARIANT intercept, VARIANT displayEquation,
                                     VARIANT displayRSquared, VARIANT name,
                                     ExcelObj_Trendline *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Trendlines, 0xB5, CAVT_OBJHANDLE,
	                              returnValue, 9, __paramTypes, type, order,
	                              period, forward, backward, intercept,
	                              displayEquation, displayRSquared, name);

	return __result;
}

HRESULT CVIFUNC Excel_TrendlinesItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Trendline *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Trendlines, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Trendlines_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Trendlines, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_TrendlineClearFormats (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Trendline, 0x70, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_TrendlineDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Trendline, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_TrendlineSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Trendline, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_CornersSelect (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Corners,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesLinesSelect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SeriesLines, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_SeriesLinesDelete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_SeriesLines, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_HiLoLinesSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_HiLoLines, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_HiLoLinesDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_HiLoLines, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_GridlinesSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Gridlines, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_GridlinesDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Gridlines, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DropLinesSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DropLines, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DropLinesDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DropLines, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LeaderLinesDelete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LeaderLines, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LeaderLinesSelect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LeaderLines, 0xEB, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_UpBarsSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_UpBars,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_UpBarsDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_UpBars,
	                              0x75, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DownBarsSelect (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_DownBars,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DownBarsDelete (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_DownBars,
	                              0x75, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_FloorSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Floor,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_FloorClearFormats (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Floor,
	                              0x70, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_FloorPaste (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Floor,
	                              0xD3, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WallsSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Walls,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WallsClearFormats (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Walls,
	                              0x70, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_WallsPaste (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Walls,
	                              0xD3, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_TickLabelsDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_TickLabels, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_TickLabelsSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_TickLabels, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PlotAreaSelect (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_PlotArea,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PlotAreaClearFormats (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_PlotArea,
	                              0x70, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartAreaSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartArea, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartAreaClear (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartArea, 0x6F, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartAreaClearContents (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartArea, 0x71, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartAreaCopy (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartArea, 0x227, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ChartAreaClearFormats (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ChartArea, 0x70, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LegendSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Legend,
	                              0xEB, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LegendDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Legend,
	                              0x75, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_LegendLegendEntries (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Legend,
	                              0xAD, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_LegendClear (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Legend,
	                              0x6F, CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ErrorBarsSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ErrorBars, 0xEB, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ErrorBarsDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ErrorBars, 0x75, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ErrorBarsClearFormats (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ErrorBars, 0x70, CAVT_VARIANT,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataTableSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataTable, 0xEB, CAVT_EMPTY, NULL,
	                              0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DataTableDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DataTable, 0x75, CAVT_EMPTY, NULL,
	                              0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeApply (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x68B, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x75, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeDuplicate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      ExcelObj_Shape *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x40F, CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeFlip (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 enum ExcelEnum_MsoFlipCmd flipCmd)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x68C, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              flipCmd);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeIncrementLeft (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x68E, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              increment);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeIncrementRotation (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x690, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              increment);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeIncrementTop (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x691, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              increment);

	return __result;
}

HRESULT CVIFUNC Excel_ShapePickUp (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x692, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRerouteConnections (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x693, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeScaleHeight (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, float factor,
                                        enum ExcelEnum_MsoTriState relativeToOriginalSize,
                                        VARIANT scale)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT, CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x694, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              factor, relativeToOriginalSize, scale);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeScaleWidth (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, float factor,
                                       enum ExcelEnum_MsoTriState relativeToOriginalSize,
                                       VARIANT scale)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT, CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x698, CAVT_EMPTY, NULL, 3, __paramTypes,
	                              factor, relativeToOriginalSize, scale);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT replace)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0xEB, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              replace);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeSetShapesDefaultProperties (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x699, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeUngroup (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo,
                                    ExcelObj_ShapeRange *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0xF4, CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeZOrder (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo,
                                   enum ExcelEnum_MsoZOrderCmd ZOrderCmd)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x26E, CAVT_EMPTY, NULL, 1, __paramTypes,
	                              ZOrderCmd);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x227, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeCut (CAObjHandle objectHandle, ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0x235, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeCopyPicture (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT appearance,
                                        VARIANT format)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shape,
	                              0xD5, CAVT_EMPTY, NULL, 2, __paramTypes,
	                              appearance, format);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0xAA, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_Shapes_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x0, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddCallout (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_MsoCalloutType type,
                                        float left, float top, float width,
                                        float height,
                                        ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_FLOAT, CAVT_FLOAT,
	                               CAVT_FLOAT, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6B1, CAVT_OBJHANDLE, returnValue, 5,
	                              __paramTypes, type, left, top, width, height);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddConnector (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          enum ExcelEnum_MsoConnectorType type,
                                          float beginX, float beginY, float endX,
                                          float endY,
                                          ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_FLOAT, CAVT_FLOAT,
	                               CAVT_FLOAT, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6B2, CAVT_OBJHANDLE, returnValue, 5,
	                              __paramTypes, type, beginX, beginY, endX,
	                              endY);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddCurve (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      VARIANT safeArrayOfPoints,
                                      ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6B7, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, safeArrayOfPoints);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddLabel (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_MsoTextOrientation orientation,
                                      float left, float top, float width,
                                      float height, ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_FLOAT, CAVT_FLOAT,
	                               CAVT_FLOAT, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6B9, CAVT_OBJHANDLE, returnValue, 5,
	                              __paramTypes, orientation, left, top, width,
	                              height);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddLine (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, float beginX,
                                     float beginY, float endX, float endY,
                                     ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT, CAVT_FLOAT, CAVT_FLOAT,
	                               CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6BA, CAVT_OBJHANDLE, returnValue, 4,
	                              __paramTypes, beginX, beginY, endX, endY);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddPicture (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        const char *filename,
                                        enum ExcelEnum_MsoTriState linkToFile,
                                        enum ExcelEnum_MsoTriState saveWithDocument,
                                        float left, float top, float width,
                                        float height,
                                        ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_LONG, CAVT_LONG,
	                               CAVT_FLOAT, CAVT_FLOAT, CAVT_FLOAT,
	                               CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6BB, CAVT_OBJHANDLE, returnValue, 7,
	                              __paramTypes, filename, linkToFile,
	                              saveWithDocument, left, top, width, height);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddPolyline (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT safeArrayOfPoints,
                                         ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6BE, CAVT_OBJHANDLE, returnValue, 1,
	                              __paramTypes, safeArrayOfPoints);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddShape (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_MsoAutoShapeType type,
                                      float left, float top, float width,
                                      float height, ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_FLOAT, CAVT_FLOAT,
	                               CAVT_FLOAT, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6BF, CAVT_OBJHANDLE, returnValue, 5,
	                              __paramTypes, type, left, top, width, height);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddTextEffect (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_MsoPresetTextEffect presetTextEffect,
                                           const char *text,
                                           const char *fontName, float fontSize,
                                           enum ExcelEnum_MsoTriState fontBold,
                                           enum ExcelEnum_MsoTriState fontItalic,
                                           float left, float top,
                                           ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_CSTRING, CAVT_CSTRING,
	                               CAVT_FLOAT, CAVT_LONG, CAVT_LONG,
	                               CAVT_FLOAT, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6C0, CAVT_OBJHANDLE, returnValue, 8,
	                              __paramTypes, presetTextEffect, text,
	                              fontName, fontSize, fontBold, fontItalic,
	                              left, top);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddTextbox (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_MsoTextOrientation orientation,
                                        float left, float top, float width,
                                        float height,
                                        ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_FLOAT, CAVT_FLOAT,
	                               CAVT_FLOAT, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6C6, CAVT_OBJHANDLE, returnValue, 5,
	                              __paramTypes, orientation, left, top, width,
	                              height);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesBuildFreeform (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_MsoEditingType editingType,
                                           float x1, float y1,
                                           ExcelObj_FreeformBuilder *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_FLOAT, CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6C7, CAVT_OBJHANDLE, returnValue, 3,
	                              __paramTypes, editingType, x1, y1);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesRange (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT index,
                                   ExcelObj_ShapeRange *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0xC5, DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesSelectAll (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6C9, CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddFormControl (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlFormControl type,
                                            long left, long top, long width,
                                            long height,
                                            ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_LONG, CAVT_LONG,
	                               CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6CA, CAVT_OBJHANDLE, returnValue, 5,
	                              __paramTypes, type, left, top, width, height);

	return __result;
}

HRESULT CVIFUNC Excel_ShapesAddOLEObject (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT classType, VARIANT filename,
                                          VARIANT link, VARIANT displayAsIcon,
                                          VARIANT iconFileName,
                                          VARIANT iconIndex, VARIANT iconLabel,
                                          VARIANT left, VARIANT top,
                                          VARIANT width, VARIANT height,
                                          ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo, &Excel_IID_Shapes,
	                              0x6CB, CAVT_OBJHANDLE, returnValue, 11,
	                              __paramTypes, classType, filename, link,
	                              displayAsIcon, iconFileName, iconIndex,
	                              iconLabel, left, top, width, height);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRange_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x0, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeAlign (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       enum ExcelEnum_MsoAlignCmd alignCmd,
                                       enum ExcelEnum_MsoTriState relativeTo)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x6CC, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, alignCmd, relativeTo);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeApply (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x68B, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeDistribute (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_MsoDistributeCmd distributeCmd,
                                            enum ExcelEnum_MsoTriState relativeTo)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x6CE, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, distributeCmd,
	                              relativeTo);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeDuplicate (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           ExcelObj_ShapeRange *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x40F, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeFlip (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_MsoFlipCmd flipCmd)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x68C, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, flipCmd);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeIncrementLeft (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x68E, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, increment);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeIncrementRotation (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x690, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, increment);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeIncrementTop (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              float increment)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x691, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, increment);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeGroup (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       ExcelObj_Shape *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x2E, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangePickUp (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x692, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeRerouteConnections (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x693, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeRegroup (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         ExcelObj_Shape *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x6D0, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeScaleHeight (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, float factor,
                                             enum ExcelEnum_MsoTriState relativeToOriginalSize,
                                             VARIANT scale)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT, CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x694, CAVT_EMPTY,
	                              NULL, 3, __paramTypes, factor,
	                              relativeToOriginalSize, scale);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeScaleWidth (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, float factor,
                                            enum ExcelEnum_MsoTriState relativeToOriginalSize,
                                            VARIANT scale)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_FLOAT, CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x698, CAVT_EMPTY,
	                              NULL, 3, __paramTypes, factor,
	                              relativeToOriginalSize, scale);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT replace)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0xEB, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, replace);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeSetShapesDefaultProps (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x699, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeUngroup (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         ExcelObj_ShapeRange *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0xF4, CAVT_OBJHANDLE,
	                              returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ShapeRangeZOrder (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_MsoZOrderCmd ZOrderCmd)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ShapeRange, 0x26E, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, ZOrderCmd);

	return __result;
}

HRESULT CVIFUNC Excel_GroupShapesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_GroupShapes, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_GroupShapes_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           ExcelObj_Shape *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_GroupShapes, 0x0, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_TextFrameCharacters (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT start,
                                           VARIANT length,
                                           ExcelObj_Characters *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_TextFrame, 0x25B, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, start, length);

	return __result;
}

HRESULT CVIFUNC Excel_ConnectorFormatBeginConnect (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   ExcelObj_Shape connectedShape,
                                                   long connectionSite)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ConnectorFormat, 0x6D6,
	                              CAVT_EMPTY, NULL, 2, __paramTypes,
	                              connectedShape, connectionSite);

	return __result;
}

HRESULT CVIFUNC Excel_ConnectorFormatBeginDisconnect (CAObjHandle objectHandle,
                                                      ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ConnectorFormat, 0x6D9,
	                              CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ConnectorFormatEndConnect (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Shape connectedShape,
                                                 long connectionSite)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_OBJHANDLE, CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ConnectorFormat, 0x6DA,
	                              CAVT_EMPTY, NULL, 2, __paramTypes,
	                              connectedShape, connectionSite);

	return __result;
}

HRESULT CVIFUNC Excel_ConnectorFormatEndDisconnect (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ConnectorFormat, 0x6DB,
	                              CAVT_EMPTY, NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_FreeformBuilderAddNodes (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               enum ExcelEnum_MsoSegmentType segmentType,
                                               enum ExcelEnum_MsoEditingType editingType,
                                               float x1, float y1, VARIANT x2,
                                               VARIANT y2, VARIANT x3,
                                               VARIANT y3)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_FLOAT,
	                               CAVT_FLOAT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FreeformBuilder, 0x6E2,
	                              CAVT_EMPTY, NULL, 8, __paramTypes,
	                              segmentType, editingType, x1, y1, x2, y2, x3,
	                              y3);

	return __result;
}

HRESULT CVIFUNC Excel_FreeformBuilderConvertToShape (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     ExcelObj_Shape *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_FreeformBuilder, 0x6E6,
	                              CAVT_OBJHANDLE, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ControlFormatAddItem (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            const char *text, VARIANT index)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_CSTRING, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ControlFormat, 0x353, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, text, index);

	return __result;
}

HRESULT CVIFUNC Excel_ControlFormatRemoveAllItems (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ControlFormat, 0x355, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_ControlFormatRemoveItem (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, long index,
                                               VARIANT count)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ControlFormat, 0x354, CAVT_EMPTY,
	                              NULL, 2, __paramTypes, index, count);

	return __result;
}

HRESULT CVIFUNC Excel_ControlFormatList (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         VARIANT *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_ControlFormat, 0x35D,
	                              CAVT_VARIANT, returnValue, 1, __paramTypes,
	                              index);

	return __result;
}

HRESULT CVIFUNC Excel_OLEFormatActivate (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEFormat, 0x130, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEFormatVerb (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT verb)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEFormat, 0x25E, CAVT_EMPTY,
	                              NULL, 1, __paramTypes, verb);

	return __result;
}

HRESULT CVIFUNC Excel_LinkFormatUpdate (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_LinkFormat, 0x2A8, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PublishObjectsAdd (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlSourceType sourceType,
                                         const char *filename, VARIANT sheet,
                                         VARIANT source, VARIANT htmlType,
                                         VARIANT divID, VARIANT title,
                                         ExcelObj_PublishObject *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_CSTRING, CAVT_VARIANT,
	                               CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PublishObjects, 0xB5,
	                              CAVT_OBJHANDLE, returnValue, 7, __paramTypes,
	                              sourceType, filename, sheet, source,
	                              htmlType, divID, title);

	return __result;
}

HRESULT CVIFUNC Excel_PublishObjectsItem (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_PublishObject *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PublishObjects, 0xAA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PublishObjects_Default (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              ExcelObj_PublishObject *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PublishObjects, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PublishObjectsDelete (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PublishObjects, 0x75, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PublishObjectsPublish (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PublishObjects, 0x767, CAVT_EMPTY,
	                              NULL, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_OLEDBErrorsItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_OLEDBError *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEDBErrors, 0xAA, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_OLEDBErrors_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long index,
                                           ExcelObj_OLEDBError *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEDBErrors, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_OLEDBErrors_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_OLEDBErrors, 0xFFFFFFFC,
	                              CAVT_UNKNOWN, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PhoneticsItem (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long index,
                                     CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Phonetics, 0xAA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PhoneticsDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Phonetics, 0x75, CAVT_EMPTY, NULL,
	                              0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_PhoneticsAdd (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, long start,
                                    long length, const char *text)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG, CAVT_LONG, CAVT_CSTRING};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_Phonetics, 0xB5, CAVT_EMPTY, NULL,
	                              3, __paramTypes, start, length, text);

	return __result;
}

HRESULT CVIFUNC Excel_Phonetics_Default (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, long index,
                                         CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_LONG};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_Phonetics, 0x0,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotLayoutColumnFields (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT index,
                                               CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotLayout, 0x2C9,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotLayoutDataFields (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotLayout, 0x2CB,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotLayoutPageFields (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotLayout, 0x2CA,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotLayoutRowFields (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotLayout, 0x2C8,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotLayoutHiddenFields (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT index,
                                               CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotLayout, 0x2C7,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotLayoutVisibleFields (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index,
                                                CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotLayout, 0x2C6,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotLayoutPivotFields (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              CAObjHandle *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotLayout, 0x2CE,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 1, __paramTypes, index);

	return __result;
}

HRESULT CVIFUNC Excel_PivotLayoutAddFields (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT rowFields,
                                            VARIANT columnFields,
                                            VARIANT pageFields,
                                            VARIANT appendField)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT, CAVT_VARIANT,
	                               CAVT_VARIANT};

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_PivotLayout, 0x2C4, CAVT_EMPTY,
	                              NULL, 4, __paramTypes, rowFields,
	                              columnFields, pageFields, appendField);

	return __result;
}

HRESULT CVIFUNC Excel_DisplayUnitLabelSelect (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DisplayUnitLabel, 0xEB,
	                              CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DisplayUnitLabelDelete (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue)
{
	HRESULT __result;

	__result = CA_MethodInvokeEx (objectHandle, errorInfo,
	                              &Excel_IID_DisplayUnitLabel, 0x75,
	                              CAVT_VARIANT, returnValue, 0, NULL);

	return __result;
}

HRESULT CVIFUNC Excel_DisplayUnitLabelCharacters (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT start, VARIANT length,
                                                  ExcelObj_Characters *returnValue)
{
	HRESULT __result;
	unsigned int __paramTypes[] = {CAVT_VARIANT, CAVT_VARIANT};

	__result = CA_InvokeHelperEx (objectHandle, errorInfo,
	                              &Excel_IID_DisplayUnitLabel, 0x25B,
	                              DISPATCH_PROPERTYGET, CAVT_OBJHANDLE,
	                              returnValue, 2, __paramTypes, start, length);

	return __result;
}

static void CVIFUNC RefreshEventsRegOnBeforeRefresh_EventVTableFunc (void *thisPtr,
                                                                     VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	RefreshEventsRegOnBeforeRefresh_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 0, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, cancel);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC RefreshEventsRegOnAfterRefresh_EventVTableFunc (void *thisPtr,
                                                                    VBOOL success)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	RefreshEventsRegOnAfterRefresh_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 1, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, success);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC AppEventsRegOnNewWorkbook_EventVTableFunc (void *thisPtr,
                                                               LPDISPATCH wb)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnNewWorkbook_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 0, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);

	return;
}

static void CVIFUNC AppEventsRegOnSheetSelectionChange_EventVTableFunc (void *thisPtr,
                                                                        LPDISPATCH sh,
                                                                        LPDISPATCH target)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnSheetSelectionChange_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 1, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC AppEventsRegOnSheetBeforeDoubleClick_EventVTableFunc (void *thisPtr,
                                                                          LPDISPATCH sh,
                                                                          LPDISPATCH target,
                                                                          VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnSheetBeforeDoubleClick_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 2, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType, cancel);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC AppEventsRegOnSheetBeforeRightClick_EventVTableFunc (void *thisPtr,
                                                                         LPDISPATCH sh,
                                                                         LPDISPATCH target,
                                                                         VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnSheetBeforeRightClick_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 3, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType, cancel);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC AppEventsRegOnSheetActivate_EventVTableFunc (void *thisPtr,
                                                                 LPDISPATCH sh)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnSheetActivate_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 4, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);

	return;
}

static void CVIFUNC AppEventsRegOnSheetDeactivate_EventVTableFunc (void *thisPtr,
                                                                   LPDISPATCH sh)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnSheetDeactivate_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 5, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);

	return;
}

static void CVIFUNC AppEventsRegOnSheetCalculate_EventVTableFunc (void *thisPtr,
                                                                  LPDISPATCH sh)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnSheetCalculate_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 6, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);

	return;
}

static void CVIFUNC AppEventsRegOnSheetChange_EventVTableFunc (void *thisPtr,
                                                               LPDISPATCH sh,
                                                               LPDISPATCH target)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnSheetChange_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 7, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWorkbookOpen_EventVTableFunc (void *thisPtr,
                                                                LPDISPATCH wb)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWorkbookOpen_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 8, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWorkbookActivate_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH wb)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWorkbookActivate_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 9, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWorkbookDeactivate_EventVTableFunc (void *thisPtr,
                                                                      LPDISPATCH wb)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWorkbookDeactivate_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 10, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWorkbookBeforeClose_EventVTableFunc (void *thisPtr,
                                                                       LPDISPATCH wb,
                                                                       VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWorkbookBeforeClose_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 11, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType,
		                               cancel);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWorkbookBeforeSave_EventVTableFunc (void *thisPtr,
                                                                      LPDISPATCH wb,
                                                                      VBOOL saveAsUI,
                                                                      VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWorkbookBeforeSave_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 12, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType,
		                               saveAsUI, cancel);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWorkbookBeforePrint_EventVTableFunc (void *thisPtr,
                                                                       LPDISPATCH wb,
                                                                       VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWorkbookBeforePrint_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 13, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType,
		                               cancel);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWorkbookNewSheet_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH wb,
                                                                    LPDISPATCH sh)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWorkbookNewSheet_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	CAObjHandle sh__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 14, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType,
		                               sh__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);
	CA_DiscardObjHandle (sh__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWorkbookAddinInstall_EventVTableFunc (void *thisPtr,
                                                                        LPDISPATCH wb)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWorkbookAddinInstall_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 15, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWorkbookAddinUninstall_EventVTableFunc (void *thisPtr,
                                                                          LPDISPATCH wb)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWorkbookAddinUninstall_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 16, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWindowResize_EventVTableFunc (void *thisPtr,
                                                                LPDISPATCH wb,
                                                                LPDISPATCH wn)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWindowResize_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	CAObjHandle wn__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 17, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (wn, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wn__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType,
		                               wn__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);
	CA_DiscardObjHandle (wn__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWindowActivate_EventVTableFunc (void *thisPtr,
                                                                  LPDISPATCH wb,
                                                                  LPDISPATCH wn)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWindowActivate_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	CAObjHandle wn__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 18, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (wn, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wn__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType,
		                               wn__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);
	CA_DiscardObjHandle (wn__CType);

	return;
}

static void CVIFUNC AppEventsRegOnWindowDeactivate_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH wb,
                                                                    LPDISPATCH wn)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnWindowDeactivate_CallbackType __callbackFunction;
	CAObjHandle wb__CType = 0;
	CAObjHandle wn__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 19, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wb, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wb__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (wn, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wn__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wb__CType,
		                               wn__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wb__CType);
	CA_DiscardObjHandle (wn__CType);

	return;
}

static void CVIFUNC AppEventsRegOnSheetFollowHyperlink_EventVTableFunc (void *thisPtr,
                                                                        LPDISPATCH sh,
                                                                        LPDISPATCH target)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	AppEventsRegOnSheetFollowHyperlink_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 20, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC ChartEventsRegOnActivate_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnActivate_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 0, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnDeactivate_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnDeactivate_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 1, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnResize_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnResize_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 2, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnMouseDown_EventVTableFunc (void *thisPtr,
                                                               long button,
                                                               long shift,
                                                               long x, long y)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnMouseDown_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 3, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, button, shift,
		                               x, y);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnMouseUp_EventVTableFunc (void *thisPtr,
                                                             long button,
                                                             long shift, long x,
                                                             long y)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnMouseUp_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 4, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, button, shift,
		                               x, y);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnMouseMove_EventVTableFunc (void *thisPtr,
                                                               long button,
                                                               long shift,
                                                               long x, long y)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnMouseMove_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 5, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, button, shift,
		                               x, y);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnBeforeRightClick_EventVTableFunc (void *thisPtr,
                                                                      VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnBeforeRightClick_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 6, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, cancel);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnDragPlot_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnDragPlot_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 7, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnDragOver_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnDragOver_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 8, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnBeforeDoubleClick_EventVTableFunc (void *thisPtr,
                                                                       long elementID,
                                                                       long arg1,
                                                                       long arg2,
                                                                       VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnBeforeDoubleClick_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 9, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, elementID,
		                               arg1, arg2, cancel);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnSelect_EventVTableFunc (void *thisPtr,
                                                            long elementID,
                                                            long arg1, long arg2)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnSelect_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 10, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, elementID,
		                               arg1, arg2);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnSeriesChange_EventVTableFunc (void *thisPtr,
                                                                  long seriesIndex,
                                                                  long pointIndex)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnSeriesChange_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 11, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, seriesIndex,
		                               pointIndex);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC ChartEventsRegOnCalculate_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	ChartEventsRegOnCalculate_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 12, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC DocEventsRegOnSelectionChange_EventVTableFunc (void *thisPtr,
                                                                   LPDISPATCH target)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	DocEventsRegOnSelectionChange_CallbackType __callbackFunction;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 0, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, target__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC DocEventsRegOnBeforeDoubleClick_EventVTableFunc (void *thisPtr,
                                                                     LPDISPATCH target,
                                                                     VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	DocEventsRegOnBeforeDoubleClick_CallbackType __callbackFunction;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 1, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, target__CType,
		                               cancel);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC DocEventsRegOnBeforeRightClick_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH target,
                                                                    VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	DocEventsRegOnBeforeRightClick_CallbackType __callbackFunction;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 2, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, target__CType,
		                               cancel);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC DocEventsRegOnActivate_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	DocEventsRegOnActivate_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 3, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC DocEventsRegOnDeactivate_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	DocEventsRegOnDeactivate_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 4, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC DocEventsRegOnCalculate_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	DocEventsRegOnCalculate_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 5, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC DocEventsRegOnChange_EventVTableFunc (void *thisPtr,
                                                          LPDISPATCH target)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	DocEventsRegOnChange_CallbackType __callbackFunction;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 6, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, target__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC DocEventsRegOnFollowHyperlink_EventVTableFunc (void *thisPtr,
                                                                   LPDISPATCH target)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	DocEventsRegOnFollowHyperlink_CallbackType __callbackFunction;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 7, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, target__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnOpen_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnOpen_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 0, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC WorkbkEventsRegOnActivate_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnActivate_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 1, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC WorkbkEventsRegOnDeactivate_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnDeactivate_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 2, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC WorkbkEventsRegOnBeforeClose_EventVTableFunc (void *thisPtr,
                                                                  VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnBeforeClose_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 3, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, cancel);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC WorkbkEventsRegOnBeforeSave_EventVTableFunc (void *thisPtr,
                                                                 VBOOL saveAsUI,
                                                                 VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnBeforeSave_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 4, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, saveAsUI,
		                               cancel);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC WorkbkEventsRegOnBeforePrint_EventVTableFunc (void *thisPtr,
                                                                  VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnBeforePrint_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 5, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, cancel);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC WorkbkEventsRegOnNewSheet_EventVTableFunc (void *thisPtr,
                                                               LPDISPATCH sh)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnNewSheet_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 6, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnAddinInstall_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnAddinInstall_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 7, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC WorkbkEventsRegOnAddinUninstall_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnAddinUninstall_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 8, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC WorkbkEventsRegOnWindowResize_EventVTableFunc (void *thisPtr,
                                                                   LPDISPATCH wn)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnWindowResize_CallbackType __callbackFunction;
	CAObjHandle wn__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 9, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wn, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wn__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wn__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wn__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnWindowActivate_EventVTableFunc (void *thisPtr,
                                                                     LPDISPATCH wn)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnWindowActivate_CallbackType __callbackFunction;
	CAObjHandle wn__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 10, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wn, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wn__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wn__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wn__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnWindowDeactivate_EventVTableFunc (void *thisPtr,
                                                                       LPDISPATCH wn)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnWindowDeactivate_CallbackType __callbackFunction;
	CAObjHandle wn__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 11, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (wn, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &wn__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, wn__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (wn__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnSheetSelectionChange_EventVTableFunc (void *thisPtr,
                                                                           LPDISPATCH sh,
                                                                           LPDISPATCH target)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnSheetSelectionChange_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 12, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnSheetBeforeDoubleClick_EventVTableFunc (void *thisPtr,
                                                                             LPDISPATCH sh,
                                                                             LPDISPATCH target,
                                                                             VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnSheetBeforeDoubleClick_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 13, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType, cancel);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnSheetBeforeRightClick_EventVTableFunc (void *thisPtr,
                                                                            LPDISPATCH sh,
                                                                            LPDISPATCH target,
                                                                            VBOOL *cancel)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnSheetBeforeRightClick_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 14, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType, cancel);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnSheetActivate_EventVTableFunc (void *thisPtr,
                                                                    LPDISPATCH sh)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnSheetActivate_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 15, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnSheetDeactivate_EventVTableFunc (void *thisPtr,
                                                                      LPDISPATCH sh)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnSheetDeactivate_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 16, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnSheetCalculate_EventVTableFunc (void *thisPtr,
                                                                     LPDISPATCH sh)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnSheetCalculate_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 17, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnSheetChange_EventVTableFunc (void *thisPtr,
                                                                  LPDISPATCH sh,
                                                                  LPDISPATCH target)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnSheetChange_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 18, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC WorkbkEventsRegOnSheetFollowHyperlink_EventVTableFunc (void *thisPtr,
                                                                           LPDISPATCH sh,
                                                                           LPDISPATCH target)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	WorkbkEventsRegOnSheetFollowHyperlink_CallbackType __callbackFunction;
	CAObjHandle sh__CType = 0;
	CAObjHandle target__CType = 0;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 19, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
		__caErrChk (CA_CreateObjHandleFromInterface (sh, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &sh__CType));
		__caErrChk (CA_CreateObjHandleFromInterface (target, &IID_IDispatch, 0,
		                                             LOCALE_NEUTRAL, 0, 1,
		                                             &target__CType));
	
		__result = __callbackFunction (__serverObjHandle, __callbackData, sh__CType,
		                               target__CType);
	
		__caErrChk (__result);
		
		}
Error:
	CA_DiscardObjHandle (sh__CType);
	CA_DiscardObjHandle (target__CType);

	return;
}

static void CVIFUNC OLEObjectEventsRegOnGotFocus_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	OLEObjectEventsRegOnGotFocus_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 0, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

static void CVIFUNC OLEObjectEventsRegOnLostFocus_EventVTableFunc (void *thisPtr)
{
	HRESULT __result;
	void * __callbackData;
	CAObjHandle __serverObjHandle;
	OLEObjectEventsRegOnLostFocus_CallbackType __callbackFunction;
	
	

	__caErrChk (CA_GetEventCallback (thisPtr, 1, &__callbackFunction,
	                                 &__callbackData, &__serverObjHandle));

	if (__callbackFunction != NULL)
		{
	
		__result = __callbackFunction (__serverObjHandle, __callbackData);
	
		__caErrChk (__result);
		
		}
Error:

	return;
}

HRESULT CVIFUNC Excel_RefreshEventsRegOnBeforeRefresh (CAObjHandle serverObject,
                                                       RefreshEventsRegOnBeforeRefresh_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_RefreshEvents_CAEventClassDefn, 0,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_RefreshEventsRegOnAfterRefresh (CAObjHandle serverObject,
                                                      RefreshEventsRegOnAfterRefresh_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_RefreshEvents_CAEventClassDefn, 1,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnNewWorkbook (CAObjHandle serverObject,
                                                 AppEventsRegOnNewWorkbook_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 0,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnSheetSelectionChange (CAObjHandle serverObject,
                                                          AppEventsRegOnSheetSelectionChange_CallbackType callbackFunction,
                                                          void *callbackData,
                                                          int enableCallbacks,
                                                          int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 1,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnSheetBeforeDoubleClick (CAObjHandle serverObject,
                                                            AppEventsRegOnSheetBeforeDoubleClick_CallbackType callbackFunction,
                                                            void *callbackData,
                                                            int enableCallbacks,
                                                            int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 2,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnSheetBeforeRightClick (CAObjHandle serverObject,
                                                           AppEventsRegOnSheetBeforeRightClick_CallbackType callbackFunction,
                                                           void *callbackData,
                                                           int enableCallbacks,
                                                           int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 3,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnSheetActivate (CAObjHandle serverObject,
                                                   AppEventsRegOnSheetActivate_CallbackType callbackFunction,
                                                   void *callbackData,
                                                   int enableCallbacks,
                                                   int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 4,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnSheetDeactivate (CAObjHandle serverObject,
                                                     AppEventsRegOnSheetDeactivate_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 5,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnSheetCalculate (CAObjHandle serverObject,
                                                    AppEventsRegOnSheetCalculate_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 6,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnSheetChange (CAObjHandle serverObject,
                                                 AppEventsRegOnSheetChange_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 7,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookOpen (CAObjHandle serverObject,
                                                  AppEventsRegOnWorkbookOpen_CallbackType callbackFunction,
                                                  void *callbackData,
                                                  int enableCallbacks,
                                                  int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 8,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookActivate (CAObjHandle serverObject,
                                                      AppEventsRegOnWorkbookActivate_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 9,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookDeactivate (CAObjHandle serverObject,
                                                        AppEventsRegOnWorkbookDeactivate_CallbackType callbackFunction,
                                                        void *callbackData,
                                                        int enableCallbacks,
                                                        int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 10,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookBeforeClose (CAObjHandle serverObject,
                                                         AppEventsRegOnWorkbookBeforeClose_CallbackType callbackFunction,
                                                         void *callbackData,
                                                         int enableCallbacks,
                                                         int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 11,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookBeforeSave (CAObjHandle serverObject,
                                                        AppEventsRegOnWorkbookBeforeSave_CallbackType callbackFunction,
                                                        void *callbackData,
                                                        int enableCallbacks,
                                                        int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 12,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookBeforePrint (CAObjHandle serverObject,
                                                         AppEventsRegOnWorkbookBeforePrint_CallbackType callbackFunction,
                                                         void *callbackData,
                                                         int enableCallbacks,
                                                         int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 13,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookNewSheet (CAObjHandle serverObject,
                                                      AppEventsRegOnWorkbookNewSheet_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 14,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookAddinInstall (CAObjHandle serverObject,
                                                          AppEventsRegOnWorkbookAddinInstall_CallbackType callbackFunction,
                                                          void *callbackData,
                                                          int enableCallbacks,
                                                          int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 15,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookAddinUninstall (CAObjHandle serverObject,
                                                            AppEventsRegOnWorkbookAddinUninstall_CallbackType callbackFunction,
                                                            void *callbackData,
                                                            int enableCallbacks,
                                                            int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 16,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWindowResize (CAObjHandle serverObject,
                                                  AppEventsRegOnWindowResize_CallbackType callbackFunction,
                                                  void *callbackData,
                                                  int enableCallbacks,
                                                  int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 17,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWindowActivate (CAObjHandle serverObject,
                                                    AppEventsRegOnWindowActivate_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 18,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnWindowDeactivate (CAObjHandle serverObject,
                                                      AppEventsRegOnWindowDeactivate_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 19,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_AppEventsRegOnSheetFollowHyperlink (CAObjHandle serverObject,
                                                          AppEventsRegOnSheetFollowHyperlink_CallbackType callbackFunction,
                                                          void *callbackData,
                                                          int enableCallbacks,
                                                          int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_AppEvents_CAEventClassDefn, 20,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnActivate (CAObjHandle serverObject,
                                                ChartEventsRegOnActivate_CallbackType callbackFunction,
                                                void *callbackData,
                                                int enableCallbacks,
                                                int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 0,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnDeactivate (CAObjHandle serverObject,
                                                  ChartEventsRegOnDeactivate_CallbackType callbackFunction,
                                                  void *callbackData,
                                                  int enableCallbacks,
                                                  int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 1,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnResize (CAObjHandle serverObject,
                                              ChartEventsRegOnResize_CallbackType callbackFunction,
                                              void *callbackData,
                                              int enableCallbacks,
                                              int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 2,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnMouseDown (CAObjHandle serverObject,
                                                 ChartEventsRegOnMouseDown_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 3,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnMouseUp (CAObjHandle serverObject,
                                               ChartEventsRegOnMouseUp_CallbackType callbackFunction,
                                               void *callbackData,
                                               int enableCallbacks,
                                               int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 4,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnMouseMove (CAObjHandle serverObject,
                                                 ChartEventsRegOnMouseMove_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 5,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnBeforeRightClick (CAObjHandle serverObject,
                                                        ChartEventsRegOnBeforeRightClick_CallbackType callbackFunction,
                                                        void *callbackData,
                                                        int enableCallbacks,
                                                        int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 6,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnDragPlot (CAObjHandle serverObject,
                                                ChartEventsRegOnDragPlot_CallbackType callbackFunction,
                                                void *callbackData,
                                                int enableCallbacks,
                                                int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 7,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnDragOver (CAObjHandle serverObject,
                                                ChartEventsRegOnDragOver_CallbackType callbackFunction,
                                                void *callbackData,
                                                int enableCallbacks,
                                                int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 8,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnBeforeDoubleClick (CAObjHandle serverObject,
                                                         ChartEventsRegOnBeforeDoubleClick_CallbackType callbackFunction,
                                                         void *callbackData,
                                                         int enableCallbacks,
                                                         int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 9,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnSelect (CAObjHandle serverObject,
                                              ChartEventsRegOnSelect_CallbackType callbackFunction,
                                              void *callbackData,
                                              int enableCallbacks,
                                              int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 10,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnSeriesChange (CAObjHandle serverObject,
                                                    ChartEventsRegOnSeriesChange_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 11,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_ChartEventsRegOnCalculate (CAObjHandle serverObject,
                                                 ChartEventsRegOnCalculate_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_ChartEvents_CAEventClassDefn, 12,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_DocEventsRegOnSelectionChange (CAObjHandle serverObject,
                                                     DocEventsRegOnSelectionChange_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_DocEvents_CAEventClassDefn, 0,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_DocEventsRegOnBeforeDoubleClick (CAObjHandle serverObject,
                                                       DocEventsRegOnBeforeDoubleClick_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_DocEvents_CAEventClassDefn, 1,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_DocEventsRegOnBeforeRightClick (CAObjHandle serverObject,
                                                      DocEventsRegOnBeforeRightClick_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_DocEvents_CAEventClassDefn, 2,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_DocEventsRegOnActivate (CAObjHandle serverObject,
                                              DocEventsRegOnActivate_CallbackType callbackFunction,
                                              void *callbackData,
                                              int enableCallbacks,
                                              int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_DocEvents_CAEventClassDefn, 3,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_DocEventsRegOnDeactivate (CAObjHandle serverObject,
                                                DocEventsRegOnDeactivate_CallbackType callbackFunction,
                                                void *callbackData,
                                                int enableCallbacks,
                                                int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_DocEvents_CAEventClassDefn, 4,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_DocEventsRegOnCalculate (CAObjHandle serverObject,
                                               DocEventsRegOnCalculate_CallbackType callbackFunction,
                                               void *callbackData,
                                               int enableCallbacks,
                                               int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_DocEvents_CAEventClassDefn, 5,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_DocEventsRegOnChange (CAObjHandle serverObject,
                                            DocEventsRegOnChange_CallbackType callbackFunction,
                                            void *callbackData,
                                            int enableCallbacks, int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_DocEvents_CAEventClassDefn, 6,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_DocEventsRegOnFollowHyperlink (CAObjHandle serverObject,
                                                     DocEventsRegOnFollowHyperlink_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_DocEvents_CAEventClassDefn, 7,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnOpen (CAObjHandle serverObject,
                                             WorkbkEventsRegOnOpen_CallbackType callbackFunction,
                                             void *callbackData,
                                             int enableCallbacks,
                                             int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 0,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnActivate (CAObjHandle serverObject,
                                                 WorkbkEventsRegOnActivate_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 1,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnDeactivate (CAObjHandle serverObject,
                                                   WorkbkEventsRegOnDeactivate_CallbackType callbackFunction,
                                                   void *callbackData,
                                                   int enableCallbacks,
                                                   int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 2,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnBeforeClose (CAObjHandle serverObject,
                                                    WorkbkEventsRegOnBeforeClose_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 3,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnBeforeSave (CAObjHandle serverObject,
                                                   WorkbkEventsRegOnBeforeSave_CallbackType callbackFunction,
                                                   void *callbackData,
                                                   int enableCallbacks,
                                                   int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 4,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnBeforePrint (CAObjHandle serverObject,
                                                    WorkbkEventsRegOnBeforePrint_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 5,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnNewSheet (CAObjHandle serverObject,
                                                 WorkbkEventsRegOnNewSheet_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 6,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnAddinInstall (CAObjHandle serverObject,
                                                     WorkbkEventsRegOnAddinInstall_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 7,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnAddinUninstall (CAObjHandle serverObject,
                                                       WorkbkEventsRegOnAddinUninstall_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 8,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnWindowResize (CAObjHandle serverObject,
                                                     WorkbkEventsRegOnWindowResize_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 9,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnWindowActivate (CAObjHandle serverObject,
                                                       WorkbkEventsRegOnWindowActivate_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 10,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnWindowDeactivate (CAObjHandle serverObject,
                                                         WorkbkEventsRegOnWindowDeactivate_CallbackType callbackFunction,
                                                         void *callbackData,
                                                         int enableCallbacks,
                                                         int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 11,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetSelectionChange (CAObjHandle serverObject,
                                                             WorkbkEventsRegOnSheetSelectionChange_CallbackType callbackFunction,
                                                             void *callbackData,
                                                             int enableCallbacks,
                                                             int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 12,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetBeforeDoubleClick (CAObjHandle serverObject,
                                                               WorkbkEventsRegOnSheetBeforeDoubleClick_CallbackType callbackFunction,
                                                               void *callbackData,
                                                               int enableCallbacks,
                                                               int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 13,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetBeforeRightClick (CAObjHandle serverObject,
                                                              WorkbkEventsRegOnSheetBeforeRightClick_CallbackType callbackFunction,
                                                              void *callbackData,
                                                              int enableCallbacks,
                                                              int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 14,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetActivate (CAObjHandle serverObject,
                                                      WorkbkEventsRegOnSheetActivate_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 15,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetDeactivate (CAObjHandle serverObject,
                                                        WorkbkEventsRegOnSheetDeactivate_CallbackType callbackFunction,
                                                        void *callbackData,
                                                        int enableCallbacks,
                                                        int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 16,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetCalculate (CAObjHandle serverObject,
                                                       WorkbkEventsRegOnSheetCalculate_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 17,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetChange (CAObjHandle serverObject,
                                                    WorkbkEventsRegOnSheetChange_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 18,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetFollowHyperlink (CAObjHandle serverObject,
                                                             WorkbkEventsRegOnSheetFollowHyperlink_CallbackType callbackFunction,
                                                             void *callbackData,
                                                             int enableCallbacks,
                                                             int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_WorkbkEvents_CAEventClassDefn, 19,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectEventsRegOnGotFocus (CAObjHandle serverObject,
                                                    OLEObjectEventsRegOnGotFocus_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_OLEObjectEvents_CAEventClassDefn, 0,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}

HRESULT CVIFUNC Excel_OLEObjectEventsRegOnLostFocus (CAObjHandle serverObject,
                                                     OLEObjectEventsRegOnLostFocus_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId)
{
	HRESULT __result;

	__result = CA_RegisterEventCallback (serverObject,
	                                     &_OLEObjectEvents_CAEventClassDefn, 1,
	                                     callbackFunction, callbackData,
	                                     enableCallbacks, callbackId);

	return __result;
}
