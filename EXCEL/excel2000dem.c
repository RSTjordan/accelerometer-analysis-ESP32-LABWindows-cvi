//----------------------------------------------------------------------------
// Example program to demostrate using ActiveX Automation instrument driver to
// control Microsoft Excel 97
//----------------------------------------------------------------------------

//----------------------------------------------------------------------------
// Includes
//----------------------------------------------------------------------------
#include <cviauto.h>
#include <utility.h>
#include <ansi_c.h>
#include <userint.h>

#include "toolbox.h"
#include "exceldem.h"
#include "excel2000.h"

//----------------------------------------------------------------------------
// Defines
//----------------------------------------------------------------------------
// errChk - Macro in toolbox.h that will goto label "Error:"
//          if error returned is less than zero.
#define caErrChk errChk  
#define APP_AUTOMATION_ERR "Error:  Microsoft Excel Automation"
#define APP_WARNING "Warning"
#define EXCEL_ARRAY_OF_CELLS "A2:H11"    
#define ROWS 10
#define COLUMNS 8
#define LAUNCHERR "\
An error occurred trying to launch Excel 2000 through its automation interface.\n\n\
Ensure that Excel is installed and that you can launch it manually. If errors\n\
persist, try to launch Excel manually and use the CONNECT button instead."

//----------------------------------------------------------------------------
// Variables
//----------------------------------------------------------------------------
static int panelHandle = 0;
static int excelLaunched = 0;
static int appVisible = 1;

static HRESULT status;

static ExcelObj_App               ExcelAppHandle = 0;       
static ExcelObj_Workbooks         ExcelWorkbooksHandle = 0; 
static ExcelObj_Workbook          ExcelWorkbookHandle = 0;  
static ExcelObj_Sheets            ExcelSheetsHandle = 0;    
static ExcelObj_Worksheet         ExcelWorksheetHandle = 0; 
static ExcelObj_Range             ExcelRangeHandle = 0;     
static ExcelObj_ChartObject       ExcelChartObjHandle = 0;
static ExcelObj_Chart             ExcelChartHandle = 0;
static ExcelObj_ChartGroup        ExcelChartsHandle = 0;


static ERRORINFO ErrorInfo;
static VARIANT MyVariant;
static LPDISPATCH MyDispatch;
static VARIANT MyCellRangeV;


//----------------------------------------------------------------------------
// Prototypes
//----------------------------------------------------------------------------
static HRESULT SaveDocument (CAObjHandle ExcelWorksheetHandle, char *fileName);
HRESULT ClearObjHandle(CAObjHandle *objHandle);

static int ShutdownExcel(void);
static void ReportAppAutomationError (HRESULT hr);
static void InitVariables(void);
static int  UpdateUIRDimming(int panel);


//----------------------------------------------------------------------------
// Main
//----------------------------------------------------------------------------
int main (int argc, char *argv[])
{
    if (InitCVIRTE (0, argv, 0) == 0)
        return -1;
        
	CA_InitActiveXThreadStyleForCurrentThread (0, COINIT_APARTMENTTHREADED);

    SetSleepPolicy (VAL_SLEEP_MORE);
        
    if ((panelHandle = LoadPanel (0, "exceldem.uir", PANEL)) < 0)
        return -1;
    // Setup
    UpdateUIRDimming(panelHandle);
    InitVariables();
    DisplayPanel (panelHandle);
    
    RunUserInterface ();
    
    // Cleanup 
    ShutdownExcel();
	DiscardPanel (panelHandle);
    
    return 0;
}


//----------------------------------------------------------------------------
// InitVariables
//----------------------------------------------------------------------------
static void InitVariables(void)
{
    // Demo path and filename
    GetCtrlVal (panelHandle, PANEL_VISIBILITY, &appVisible);
    
    return;    
}    


//----------------------------------------------------------------------------
// UIR Callbacks
//----------------------------------------------------------------------------
// LaunchApp
//----------------------------------------------------------------------------
int CVICALLBACK LaunchApp (int panel, int control, int event,void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;
    
    switch (event) {
        case EVENT_COMMIT:
            // Launch App
            SetWaitCursor (1);
            error = Excel_NewApp (NULL, 1, LOCALE_NEUTRAL, 0, &ExcelAppHandle);
            SetWaitCursor (0);
            if (error<0) 
			{
        		MessagePopup (APP_AUTOMATION_ERR, LAUNCHERR);
				error = 0;
                goto Error;
			}
			
            // Make App Visible
            error = Excel_SetProperty (ExcelAppHandle, NULL, Excel_AppVisible, CAVT_BOOL, appVisible?VTRUE:VFALSE);
            if (error<0) 
                goto Error;
    
            UpdateUIRDimming(panelHandle);
            MakeApplicationActive ();
            excelLaunched = 1;
            break;
    }
    
Error:    
    if (error < 0) 
        ReportAppAutomationError (error);
        
    return 0;
}

//----------------------------------------------------------------------------
// ConnectApp
//----------------------------------------------------------------------------
int CVICALLBACK ConnectApp (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;
    switch (event)
        {
        case EVENT_COMMIT:
            // Launch App
            // Connect to existing application if available
            SetWaitCursor (1);
            error = Excel_ActiveApp (NULL, 1, LOCALE_NEUTRAL, 0, &ExcelAppHandle);
            SetWaitCursor (0);
            if (error<0) 
                goto Error;
    
            // Make App Visible
            error = Excel_SetProperty (ExcelAppHandle, NULL, Excel_AppVisible, CAVT_BOOL, appVisible?VTRUE:VFALSE);
            if (error<0) 
                goto Error;
            
            UpdateUIRDimming(panelHandle);
            MakeApplicationActive ();
            excelLaunched = 0;
            break;
        }
    return 0;   
Error:    
    if (error < 0) 
        ReportAppAutomationError (error);
        
    return 0;
}

//----------------------------------------------------------------------------
// ShutdownApp
//----------------------------------------------------------------------------
int CVICALLBACK ShutdownApp (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    switch (event) {
        case EVENT_COMMIT:
            SetWaitCursor (1);
            ShutdownExcel();  
            SetWaitCursor (0);
                
            UpdateUIRDimming(panel);
            break;
    }
    return 0;
}


//----------------------------------------------------------------------------
// ShutdownExcel
//----------------------------------------------------------------------------
static int ShutdownExcel(void) 
{
    HRESULT error = 0;

    ClearObjHandle (&ExcelRangeHandle);
    ClearObjHandle (&ExcelWorksheetHandle);
    ClearObjHandle (&ExcelSheetsHandle);
    
    if (ExcelWorkbookHandle) 
    {
        // Close workbook without saving
        error = Excel_WorkbookClose (ExcelWorkbookHandle, NULL, CA_VariantBool (VFALSE), 
            CA_DEFAULT_VAL, CA_VariantBool (VFALSE));
        if (error < 0)
            goto Error;
        
        ClearObjHandle (&ExcelWorkbookHandle);
    }
    
    ClearObjHandle (&ExcelWorkbooksHandle);
        
    if (ExcelAppHandle)
    {   
        if (excelLaunched) 
        {
            // Quit the Application
            error = Excel_AppQuit (ExcelAppHandle, &ErrorInfo);
            if (error < 0) goto Error;
        }
        
        ClearObjHandle (&ExcelAppHandle);
    } 
    
    return 0;   
Error:    
    if (error < 0)
        ReportAppAutomationError (error);
        
    return error;                    
}


//----------------------------------------------------------------------------
// ChangeVisibility
//----------------------------------------------------------------------------
int CVICALLBACK ChangeVisibility (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;

    switch (event) {
        case EVENT_COMMIT:
            GetCtrlVal(panel, control, &appVisible);
            
            if (ExcelAppHandle)
            {
                error = Excel_SetProperty (ExcelAppHandle, NULL, Excel_AppVisible, 
                    CAVT_BOOL, appVisible?VTRUE:VFALSE);
                if (error < 0)
                    ReportAppAutomationError (error);
            }
            break;
    }
    return 0;
}


//----------------------------------------------------------------------------
// NewAppFile
//----------------------------------------------------------------------------
int CVICALLBACK NewAppFile (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;
    
    switch (event) {
        case EVENT_COMMIT:
            if (!ExcelWorkbooksHandle)
            {
                // Get Workbooks    
                error = Excel_GetProperty (ExcelAppHandle, NULL, Excel_AppWorkbooks, 
                    CAVT_OBJHANDLE, &ExcelWorkbooksHandle);
                if (error<0) 
                    goto Error;
    
                // Add Workbook and make it active - by default, 3 sheets will be created
                error = Excel_WorkbooksAdd (ExcelWorkbooksHandle, NULL, CA_DEFAULT_VAL, 
                    &ExcelWorkbookHandle);
                if (error<0) 
                    goto Error;

                // Get Active Workbook Sheets
                error = Excel_GetProperty (ExcelAppHandle, NULL, Excel_AppSheets, 
                    CAVT_OBJHANDLE, &ExcelSheetsHandle);
                if (error<0) 
                    goto Error;
    
                // Get First Sheet
                error = Excel_SheetsItem (ExcelSheetsHandle, NULL, CA_VariantInt(1), 
                    &ExcelWorksheetHandle);
                if (error<0) 
                    goto Error;
    
                // Make First Sheet Active - should already be active    
				error = Excel_WorksheetActivate (ExcelWorksheetHandle, NULL);
                if (error<0) 
                    goto Error;
                
                // Update UIR    
                UpdateUIRDimming(panel);
            }                                  
            else 
                MessagePopup(APP_WARNING, "Document already open");

            break;
    }
    
Error:
    ClearObjHandle (&ExcelWorksheetHandle);
    ClearObjHandle (&ExcelSheetsHandle);
    ClearObjHandle (&ExcelWorkbookHandle);
    ClearObjHandle (&ExcelWorkbooksHandle);
        
    if (error < 0) 
        ReportAppAutomationError (error);
    
    return 0;
}


//----------------------------------------------------------------------------
// OpenAppFile
//----------------------------------------------------------------------------
int CVICALLBACK OpenAppFile (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;
    char fileName[MAX_PATHNAME_LEN];
    
    switch (event) {
        case EVENT_COMMIT:
            if (!ExcelWorkbooksHandle)
            {
                // Get Workbooks    
                error = Excel_GetProperty (ExcelAppHandle, NULL, Excel_AppWorkbooks, 
                    CAVT_OBJHANDLE, &ExcelWorkbooksHandle);
                if (error<0) 
                    goto Error;
    
                // Open existing Workbook
                GetProjectDir (fileName);
                strcat(fileName, "\\exceldem.xls");
                error = Excel_WorkbooksOpen (ExcelWorkbooksHandle, NULL, fileName, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                             CA_DEFAULT_VAL, &ExcelWorkbookHandle);
                if (error<0) 
                    goto Error;

                // Get Active Workbook Sheets
                error = Excel_GetProperty (ExcelAppHandle, NULL, Excel_AppSheets, 
                    CAVT_OBJHANDLE, &ExcelSheetsHandle);
                if (error<0) 
                    goto Error;
    
                // Get First Sheet
                error = Excel_SheetsItem (ExcelSheetsHandle, NULL, CA_VariantInt(1), 
                    &ExcelWorksheetHandle);
                if (error<0) 
                    goto Error;
    
                // Make First Sheet Active - should already be active    
				error = Excel_WorksheetActivate (ExcelWorksheetHandle, NULL);
                if (error<0) 
                    goto Error;
                
                // Update UIR    
                UpdateUIRDimming(panel);
            }                                  
            else 
                MessagePopup(APP_WARNING, "Document already open");

            break;
    }
    
    return 0;    
Error:
    ClearObjHandle (&ExcelWorksheetHandle);
    ClearObjHandle (&ExcelSheetsHandle);
    ClearObjHandle (&ExcelWorkbookHandle);
    ClearObjHandle (&ExcelWorkbooksHandle);
        
    if (error < 0) 
        ReportAppAutomationError (error);
    
    return 0;
}



//----------------------------------------------------------------------------
// PrintAppFile
//----------------------------------------------------------------------------
int CVICALLBACK PrintAppFile (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;
    
    switch (event) {
        case EVENT_COMMIT:
            if (ExcelWorksheetHandle) 
            {
                error = Excel_Worksheet_PrintOut (ExcelWorksheetHandle, NULL, CA_DEFAULT_VAL,
                                 CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                 CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                 CA_DEFAULT_VAL, CA_DEFAULT_VAL);
                if (error<0) 
                    goto Error;
                                   
                MessagePopup("Print", "Document printed...");
            }
            break;
    }
    return 0;   
Error:    
    if (error < 0) 
        ReportAppAutomationError (error);
    return 0;
}


//----------------------------------------------------------------------------
// SaveAppFile
//----------------------------------------------------------------------------
int CVICALLBACK SaveAppFile (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;
    char fileName[MAX_PATHNAME_LEN];
    
    
    switch (event) {
        case EVENT_COMMIT:
            if (ExcelWorkbookHandle) 
            { 
                GetProjectDir (fileName);
                if (FileSelectPopupEx (fileName, "*.xls", "*.xls",
                                          "Save file as...", VAL_SAVE_BUTTON,
                                          0, 1, fileName)>0)
                {                      
                    SetWaitCursor (1);
                    
                    error = CA_VariantSetCString(&MyVariant, fileName);
                    error = Excel_WorkbookSaveAs (ExcelWorkbookHandle, NULL, MyVariant,
                                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                                  CA_DEFAULT_VAL, ExcelConst_xlNoChange,
                                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL);
                    CA_VariantClear(&MyVariant);
                    SetWaitCursor (0);
                    
                    if (error < 0)
                        goto Error;
                }       
            }        
            break;
    }
    return 0;   
Error:    
    if (error < 0) 
        ReportAppAutomationError (error);
    return 0;
}

//----------------------------------------------------------------------------
// CloseAppFile
//----------------------------------------------------------------------------
int CVICALLBACK CloseAppFile (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;
    
    switch (event) {
        case EVENT_COMMIT:
                
                if (ExcelWorkbookHandle) 
                {
                    ClearObjHandle (&ExcelRangeHandle);
                    ClearObjHandle (&ExcelSheetsHandle);
                    ClearObjHandle (&ExcelWorksheetHandle);
    
                    if (ExcelWorkbookHandle) 
                    {
                        // Close workbook without saving
                        error = Excel_WorkbookClose (ExcelWorkbookHandle, NULL, CA_VariantBool (VFALSE), 
                            CA_DEFAULT_VAL, CA_VariantBool (VFALSE));
                        if (error < 0)
                            goto Error;
        
                        ClearObjHandle (&ExcelWorkbookHandle);
                    }
    
                    ClearObjHandle (&ExcelWorkbooksHandle);
        
                    ExcelWorksheetHandle = 0;
                    UpdateUIRDimming(panel);
                }
            break;
    }
    return 0;   
Error:    
    if (error < 0) 
        ReportAppAutomationError (error);
    return 0;
}


//----------------------------------------------------------------------------
// WriteDataToExcel
//----------------------------------------------------------------------------
// This function illustrates 3 different ways of writing data to the Excel cells.
//----------------------------------------------------------------------------
HRESULT WriteDataToExcel(void)
{
    VARIANT *vArray = NULL;
    LPSAFEARRAY MySafeArray = NULL;
    HRESULT error = 0;
    int i, j;

    SetWaitCursor (1);
    
    // Open new Range for Worksheet
    error = CA_VariantSetCString (&MyCellRangeV, EXCEL_ARRAY_OF_CELLS);
    error = Excel_WorksheetRange (ExcelWorksheetHandle, NULL, MyCellRangeV, CA_DEFAULT_VAL, &ExcelRangeHandle);
    if (error<0) goto Error;

    // Make range Active    
    error = Excel_RangeActivate (ExcelRangeHandle, &ErrorInfo, NULL);
    if (error<0) goto Error;

    //----------------------------------------------------------------
    // 1) Set all cells in Range to a single value of zero
    //----------------------------------------------------------------
    error = Excel_SetProperty (ExcelRangeHandle, &ErrorInfo, Excel_RangeValue2, CAVT_VARIANT, CA_VariantFloat (0.0));
    if (error<0) goto Error;

    
    //----------------------------------------------------------------
    // 2) Set each cell in Range one at a time using an offset from 
    //    range's top left cell
    //----------------------------------------------------------------
    for (i=0;i<ROWS;i++)
    {
        for (j=0;j<COLUMNS;j++)
        {
            error = Excel_RangeSetItem (ExcelRangeHandle, &ErrorInfo, CA_VariantInt (i+1), CA_VariantInt (j+1), CA_VariantFloat ( sin(3.14*(i+1)/(ROWS+1)) * sin(3.14*(j+1)/(COLUMNS+1))) );
            if (error<0) goto Error;
        }
    }    

    //----------------------------------------------------------------
    // 3) Set all cells at once using a SAFEARRAY in a VARIANT
	// NOTE: The arrays must be 2-dimensional even if setting only 
	//       one row or one column.
    //----------------------------------------------------------------
    // Create a Variant Array and set each value
    vArray = (VARIANT *) calloc (ROWS*COLUMNS, sizeof(VARIANT));
    if (!vArray)
        goto Error;
    
    for (i=0;i<ROWS;i++)
    {
        for (j=0;j<COLUMNS;j++)
        {
            error = CA_VariantSetDouble (&vArray[i*COLUMNS+j], sin(3.14*(i+1)/(ROWS+1)) * sin(3.14*(j+1)/(COLUMNS+1)));
            if (error<0) goto Error;
        }
    }  
    
    // Create a SAFEARRAY
    error = CA_Array2DToSafeArray (vArray, CAVT_VARIANT, ROWS, COLUMNS, &MySafeArray);
    if (error<0) goto Error;
    
    // Set SafeArray into a Variant to send to Excel
    error = CA_VariantSetSafeArray (&MyVariant, CAVT_VARIANT, MySafeArray);
    if (error<0) goto Error;
    
    // Set Range with one call passing SAFEARRAY as Variant
    error = Excel_SetProperty (ExcelRangeHandle, &ErrorInfo, Excel_RangeValue2, CAVT_VARIANT, MyVariant);
    if (error<0) goto Error;
            
Error:
    SetWaitCursor (0);
    
    // Free array of VARIANT
    if (vArray) 
    {
        for (i=0;i<ROWS;i++)
        {
            for (j=0;j<COLUMNS;j++)
            {
                CA_VariantClear (&vArray[i*COLUMNS+j]);
            }
        }    
        free(vArray);
    }            
    
    // Free SAFEARRAY in VARIANT        
    CA_VariantClear(&MyVariant);
    CA_VariantClear(&MyCellRangeV);
    
    // Clear Range Handle
    ClearObjHandle (&ExcelRangeHandle);

    if (error < 0) 
        ReportAppAutomationError (error);
        
    return error;
}

int CVICALLBACK WriteData (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    switch (event) 
    {
        case EVENT_COMMIT:
            WriteDataToExcel();        
            break;
    }
    return 0;
}

//----------------------------------------------------------------------------
// ReadData
//----------------------------------------------------------------------------
// This function illustrates 2 different ways of reading data from Excel cells.
//----------------------------------------------------------------------------
HRESULT ReadDataFromExcel(void)
{
    HRESULT error = 0;
    int i, j;
    VARIANT *vArray = NULL;
    size_t dim1Size, dim2Size;
    double d;
    ExcelObj_Range ExcelSingleCellRangeHandle = 0;

    SetWaitCursor (1);
    
    // Open new Range for Worksheet
    error = CA_VariantSetCString (&MyCellRangeV, EXCEL_ARRAY_OF_CELLS);
    error = Excel_WorksheetRange (ExcelWorksheetHandle, NULL, MyCellRangeV, CA_DEFAULT_VAL, &ExcelRangeHandle);
    CA_VariantClear(&MyCellRangeV);
    if (error<0) goto Error;

    // Make range Active    
    error = Excel_RangeActivate (ExcelRangeHandle, &ErrorInfo, NULL);
    if (error<0) goto Error;

    //----------------------------------------------------------------
    // 1) Get each cell value in Range one at a time using an offset 
    //    from the range's top left cell
    //----------------------------------------------------------------
    SetStdioWindowVisibility (1);
    printf("Get one cell value at a time:\n");
    for (i=0;i<ROWS;i++)
    {
        printf("    ");
        for (j=0;j<COLUMNS;j++)
        {
            // Ask for the ith by jth value of the range which returns a dispatch to a new single cell range
            error = Excel_RangeGetItem (ExcelRangeHandle, &ErrorInfo, CA_VariantInt (i+1), CA_VariantInt (j+1), &MyVariant);
            if (error<0) goto Error;
            
            // Get the DISPATCH pointer
            error = CA_VariantGetDispatch (&MyVariant, &MyDispatch);
            if (error<0) goto Error;
            
            // Create a new Range Object from DISPATCH pointer
            error = CA_CreateObjHandleFromIDispatch (MyDispatch, 0, &ExcelSingleCellRangeHandle);
            if (error<0) goto Error;
            
            // Get the value of the Single Cell Range
            error = Excel_GetProperty (ExcelSingleCellRangeHandle, &ErrorInfo, Excel_RangeValue2, CAVT_VARIANT, &MyVariant);
            if (error<0) goto Error;
            
            if (!CA_VariantHasDouble (&MyVariant))
            {
                MessagePopup(APP_WARNING, "Values returned were not of type DOUBLE.");
                goto Error;
            }    
            
            error = CA_VariantGetDouble (&MyVariant, &d);
            if (error<0) goto Error;
        
            // Free Variant element
            CA_VariantClear(&MyVariant);
    		
    		//Free Range Handle
    		ClearObjHandle (&ExcelSingleCellRangeHandle);
            
            printf("%f ", d);
        }
        printf("\n");
    }    
    printf("\n");

    //----------------------------------------------------------------
    // 2) Get Range as a SAFEARRAY inside a VARIANT
    //----------------------------------------------------------------
    error = Excel_GetProperty (ExcelRangeHandle, &ErrorInfo, Excel_RangeValue2, CAVT_VARIANT, &MyVariant);
    if (error<0) goto Error;
    
    // Get 2D Array of values from SAFEARRAY in Variant
    error = CA_VariantGet2DArray (&MyVariant, CAVT_VARIANT, &vArray, &dim1Size, &dim2Size);
    if (error<0) goto Error;
    
    // Loop on SAFEARRAY of VARIANTs
    printf("Get all data at once:\n");
    for (i = 0; i < dim1Size; i++)
    {
        printf("    ");
        for (j = 0; j < dim2Size; j++)
        {
            // Use CA_Get2DArrayElement macro to get VARIANT array element
            MyVariant = CA_Get2DArrayElement(vArray, dim1Size, dim2Size, i, j, VARIANT);
            if (!CA_VariantHasDouble (&MyVariant))
            {
                MessagePopup(APP_WARNING, "Values returned were not of type DOUBLE.");
                goto Error;
            }    
            
            // Get floating point value
            error = CA_VariantGetDouble (&MyVariant, &d);
            if (error<0) goto Error;
        
            // Clear VARAINT element in array
            CA_VariantClear(&MyVariant);

            printf("%f ", d);
        }
        printf("\n");
    }    
    printf("\n");

    
Error:
    SetWaitCursor (0);
    
    // Clear VARIANT    
    CA_VariantClear(&MyVariant);
    CA_VariantClear(&MyCellRangeV);
    
    // Free array of VARAINT
    if (vArray)
        CA_FreeMemory(vArray);
        
    // Free Range handles
    ClearObjHandle (&ExcelRangeHandle);
    ClearObjHandle (&ExcelSingleCellRangeHandle);
    
    if (error < 0) 
        ReportAppAutomationError (error);

    return error;
}


int CVICALLBACK ReadData (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    switch (event) 
    {
        case EVENT_COMMIT:
            ReadDataFromExcel();
            break;
    }
    
    return 0;
}


//----------------------------------------------------------------------------
// MakeChartInExcel
//----------------------------------------------------------------------------
HRESULT MakeChartInExcel(void)
{
    HRESULT error = 0;
    
    SetWaitCursor (1);
    
    // Open new Range for Worksheet
    error = CA_VariantSetCString (&MyCellRangeV, EXCEL_ARRAY_OF_CELLS);
    error = Excel_WorksheetRange (ExcelWorksheetHandle, NULL, MyCellRangeV, CA_DEFAULT_VAL, &ExcelRangeHandle);
    CA_VariantClear(&MyCellRangeV);
    if (error<0) goto Error;
    
    // Open new Chart Collection for Worksheet
    status = Excel_WorksheetChartObjects (ExcelWorksheetHandle, NULL, CA_DEFAULT_VAL, &ExcelChartsHandle);
    if (status<0) goto Error;
    
    // Create new chart
    status = Excel_ChartObjectsAdd (ExcelChartsHandle, NULL, 175.0, 175.0,
                                    300.0, 200.0, &ExcelChartObjHandle);
    if (status<0) goto Error;


    status = Excel_GetProperty (ExcelChartObjHandle, NULL, Excel_ChartObjectChart, CAVT_OBJHANDLE, &ExcelChartHandle);
    if (status<0) goto Error;

    // Use Chart Wizard to setup Chart
    status = CA_VariantSetCString (&MyVariant, "Chart #1");
    status = CA_GetDispatchFromObjHandle (ExcelRangeHandle, &MyDispatch);  // Get dispatch for range
    status = Excel_ChartChartWizard (ExcelChartHandle, &ErrorInfo,
                                     CA_VariantDispatch (MyDispatch),
                                     CA_VariantLong(ExcelConst_xl3DSurface),
                                     CA_DEFAULT_VAL,
                                     CA_VariantInt(ExcelConst_xlRows),
                                     CA_DEFAULT_VAL,
                                     CA_DEFAULT_VAL,
                                     CA_DEFAULT_VAL, 
                                     MyVariant,
                                     CA_DEFAULT_VAL,
                                     CA_DEFAULT_VAL,
                                     CA_DEFAULT_VAL);
    CA_VariantClear(&MyVariant);
    if (status<0) goto Error;

    // Lets get the current rotation value
    status = Excel_GetProperty (ExcelChartHandle, NULL, Excel_ChartRotation, CAVT_VARIANT, &MyVariant);
    if (status<0) goto Error;
    CA_VariantClear(&MyVariant);

Error:
    SetWaitCursor (0);
    CA_VariantClear(&MyCellRangeV);
    CA_VariantClear(&MyVariant);
    ClearObjHandle (&ExcelRangeHandle);
    ClearObjHandle (&ExcelChartHandle);
    ClearObjHandle (&ExcelChartObjHandle);
    ClearObjHandle (&ExcelChartsHandle);
    
    if (error < 0) 
        ReportAppAutomationError (error);

    return error;
}


int CVICALLBACK MakeChart (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    switch (event)
        {
        case EVENT_COMMIT:
            MakeChartInExcel();
            break;
        }
    
    return 0;
    
}


//----------------------------------------------------------------------------
// RunMacro
//----------------------------------------------------------------------------
int CVICALLBACK RunMacro (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;
    
    switch (event) 
    {
        case EVENT_COMMIT:
            SetWaitCursor (1);
            
            // Application.Run "exceldem.xls!NewDataFromMacro"
            status = CA_VariantSetCString (&MyVariant, "exceldem.xls!NewDataFromMacro");
            error = Excel_AppRun (ExcelAppHandle, NULL, MyVariant, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, CA_DEFAULT_VAL,
                                  CA_DEFAULT_VAL, CA_DEFAULT_VAL, NULL);
            CA_VariantClear(&MyVariant);
            if (status<0) goto Error;
                        
            SetWaitCursor (0);
            break;
    }
    
Error:    
    SetWaitCursor (0);
    
    if (error < 0) 
        ReportAppAutomationError (error);

    return 0;
}


//----------------------------------------------------------------------------
// Quit
//----------------------------------------------------------------------------
int CVICALLBACK Quit (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    switch (event) {
        case EVENT_COMMIT:
            QuitUserInterface (0);
            break;
    }
    return 0;
}

//----------------------------------------------------------------------------
// UpdateUIRDimming
//----------------------------------------------------------------------------
static int UpdateUIRDimming(int panel)
{
    SetCtrlAttribute (panel, PANEL_LAUNCHAPP,    ATTR_DIMMED,  (int)ExcelAppHandle);
    SetCtrlAttribute (panel, PANEL_CONNECTAPP,   ATTR_DIMMED,  (int)ExcelAppHandle);
    SetCtrlAttribute (panel, PANEL_SHUTDOWNAPP,  ATTR_DIMMED, !(int)ExcelAppHandle);
    
    SetCtrlAttribute (panel, PANEL_OPENFILE,     ATTR_DIMMED,  ((int)ExcelWorksheetHandle || !(int)ExcelAppHandle));
    SetCtrlAttribute (panel, PANEL_PRINTFILE,    ATTR_DIMMED, !(int)ExcelWorksheetHandle);
    SetCtrlAttribute (panel, PANEL_SAVEFILE,     ATTR_DIMMED, !(int)ExcelWorksheetHandle);
    SetCtrlAttribute (panel, PANEL_CLOSEFILE,    ATTR_DIMMED, !(int)ExcelWorksheetHandle);
    
    SetCtrlAttribute (panel, PANEL_WRITEDATA,    ATTR_DIMMED, !(int)ExcelWorksheetHandle);
    SetCtrlAttribute (panel, PANEL_READDATA,     ATTR_DIMMED, !(int)ExcelWorksheetHandle);
    SetCtrlAttribute (panel, PANEL_MAKECHART,    ATTR_DIMMED, !(int)ExcelWorksheetHandle);
    SetCtrlAttribute (panel, PANEL_RUNMACRO,     ATTR_DIMMED, !(int)ExcelWorksheetHandle);
    
    return 0;
}



//----------------------------------------------------------------------------
// ClearObjHandle
//----------------------------------------------------------------------------
HRESULT ClearObjHandle(CAObjHandle *objHandle)
{
    HRESULT error = 0;
    if ((objHandle) && (*objHandle))
    {
        error = CA_DiscardObjHandle (*objHandle);
        *objHandle = 0;
    }
    return error;    
}    


//----------------------------------------------------------------------------
// ReportWordAutomationError
//----------------------------------------------------------------------------
static void ReportAppAutomationError (HRESULT hr)
{
    char errorBuf[256];
    
    if (hr < 0) {
        CA_GetAutomationErrorString (hr, errorBuf, sizeof (errorBuf));
        MessagePopup (APP_AUTOMATION_ERR, errorBuf);
    }
    return;
}

