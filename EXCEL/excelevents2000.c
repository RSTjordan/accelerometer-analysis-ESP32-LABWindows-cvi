//------------------------------------------------------------------------------
// Example to demonstrate how to handle ActiveX Automation Events with LW/CVI
//------------------------------------------------------------------------------

//-----------------------------------------------------------------------------
// Includes
//-----------------------------------------------------------------------------
#include <cvirte.h>     
#include <userint.h>
#define GetSystemTime _sdk_GetSystemTime    // We want CVI's GetSystemTime
#include <cviauto.h>
#undef GetSystemTime
#include <utility.h>
#include <ansi_c.h>
#include "excel2000.h"
#include "excelevents.h"

#define APP_AUTOMATION_ERR "Error:  Microsoft Excel Automation"
#define LAUNCHERR "\
An error occurred trying to launch Excel 2000 through its automation interface.\n\n\
Ensure that Excel is installed and that you can launch it manually."

//-----------------------------------------------------------------------------
// Global Variables
//-----------------------------------------------------------------------------
static int          panelHandle = 0;
static CAObjHandle  excelApplication = 0;
static int          callbackID = 0;

//-----------------------------------------------------------------------------
// Prototypes
//-----------------------------------------------------------------------------
static void SetupButtons (int excelLaunched, int eventHandled);
static HRESULT CVICALLBACK OnBeforeCloseBook (CAObjHandle caServerObjHandle,
                                              void *caCallbackData,
                                              ExcelObj_Workbook wb,
                                              VBOOL *cancel);
//-----------------------------------------------------------------------------
// Main
//-----------------------------------------------------------------------------
int main (int argc, char *argv[])
{
    if (InitCVIRTE (0, argv, 0) == 0)
        return -1;  /* out of memory */
        
    // Force the callbacks to be called in the main thread.
    CA_InitActiveXThreadStyleForCurrentThread (0, COINIT_APARTMENTTHREADED);
    
    // Load, setup, and display panel
    if ((panelHandle = LoadPanel (0, "excelevents.uir", PANEL)) < 0)
        return -1;
    SetupButtons (0, 0);
    DisplayPanel (panelHandle);
    
    // Wait for user input
    RunUserInterface ();
    
    DiscardPanel (panelHandle);
    return 0;
}

//-----------------------------------------------------------------------------
// LaunchExcel
//-----------------------------------------------------------------------------
int CVICALLBACK LaunchExcel (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
	HRESULT error = 0;
    CAObjHandle workBooks, workBook;
    
    switch (event)
        {
        case EVENT_COMMIT:
            // Create the Excel application object
            error = Excel_NewApp (NULL, 1, LOCALE_NEUTRAL, 0, &excelApplication);
            if (error<0) 
			{
        		MessagePopup (APP_AUTOMATION_ERR, LAUNCHERR);
                goto Error;
			}
            
            // Get the Workbooks object
            Excel_GetProperty (excelApplication, NULL, Excel_AppWorkbooks,
                               CAVT_OBJHANDLE, &workBooks);

            // Show the Excel application window.  Once we make the window
            // visible, Excel will not actually shut down until we release
            // our references to its objects and the user closes the window
            Excel_SetProperty (excelApplication, NULL, Excel_AppVisible,
                               CAVT_BOOL, VTRUE);

            // Relinquish references to the workbooks so that the
            // user has control over their lifetimes through the Excel
            // application UI.
            Excel_WorkbooksAdd (workBooks, NULL, CA_DEFAULT_VAL, &workBook);
            CA_DiscardObjHandle (workBook);
            Excel_WorkbooksAdd (workBooks, NULL, CA_DEFAULT_VAL, &workBook);
            CA_DiscardObjHandle (workBook);
            Excel_WorkbooksAdd (workBooks, NULL, CA_DEFAULT_VAL, &workBook);
            CA_DiscardObjHandle (workBook);
            CA_DiscardObjHandle (workBooks);
            
            SetupButtons (1, 0);
            break;
        }
	
Error:
    return 0;
}

//-----------------------------------------------------------------------------
// ReleaseExcel
//-----------------------------------------------------------------------------
int CVICALLBACK ReleaseExcel (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    switch (event)
        {
        case EVENT_COMMIT:
            // Release the reference to the application object
            CA_DiscardObjHandle (excelApplication);
            excelApplication = 0;
            SetupButtons (0, 0);
            break;
        }
    return 0;
}

//-----------------------------------------------------------------------------
// HandleCloseBook
//-----------------------------------------------------------------------------
int CVICALLBACK HandleCloseBook (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    switch (event)
        {
        case EVENT_COMMIT:
            // Register a callback to be called before a Workbook is closed.
            Excel_AppEventsRegOnWorkbookBeforeClose (excelApplication,
                                                     OnBeforeCloseBook,
                                                     NULL, 1, &callbackID);
            SetupButtons (1, 1);
            break;
        }
    return 0;
}

//-----------------------------------------------------------------------------
// DetachFromCloseBook
//-----------------------------------------------------------------------------
int CVICALLBACK DetachFromCloseBook (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    switch (event)
        {
        case EVENT_COMMIT:
            // Unregister the callback
            CA_UnregisterEventCallback (callbackID);
            SetupButtons (1, 0);
            break;
        }
    return 0;
}

//-----------------------------------------------------------------------------
// Quit
//-----------------------------------------------------------------------------
int CVICALLBACK Quit (int panel, int control, int event,
        void *callbackData, int eventData1, int eventData2)
{
    switch (event)
        {
        case EVENT_COMMIT:
            // Release the reference to the application object
            if (excelApplication != 0)
                CA_DiscardObjHandle (excelApplication); 
            QuitUserInterface (0);
            break;
        }
    return 0;
}

//-----------------------------------------------------------------------------
// SetupButtons
//-----------------------------------------------------------------------------
// Dim the buttons to control what actions the user can perform when this 
// program is in a given state.
//-----------------------------------------------------------------------------
static void SetupButtons (int excelLaunched, int eventHandled)
{
    SetCtrlAttribute (panelHandle, PANEL_LAUNCHBUTTON, ATTR_DIMMED, excelLaunched);
    SetCtrlAttribute (panelHandle, PANEL_CLOSEBUTTON, ATTR_DIMMED, !excelLaunched);
    SetCtrlAttribute (panelHandle, PANEL_HANDLECLOSEBOOKBUTTON,
                      ATTR_DIMMED, !excelLaunched || eventHandled);
    SetCtrlAttribute (panelHandle, PANEL_DETACHCLOSEBOOKBUTON,
                      ATTR_DIMMED, !excelLaunched || !eventHandled);
}

//-----------------------------------------------------------------------------
// OnBeforeCloseBook
//-----------------------------------------------------------------------------
// This callback is called by the Excel application object when the user 
// closes a workbook.  Cancel is an output parameter.  If it is set to VTRUE,
// the Excel application object aborts the Workbook close operation.
//-----------------------------------------------------------------------------
static HRESULT CVICALLBACK OnBeforeCloseBook (CAObjHandle caServerObjHandle,
                                              void *caCallbackData,
                                              ExcelObj_Workbook wb,
                                              VBOOL *cancel)
{
    int hr, min, sec;
    int doCancel;
    char buffer[256];
    
    // Determine whether to abort the Workbook close operation
    GetCtrlVal (panelHandle, PANEL_CANCELCLOSEBOOKSWITCH, &doCancel);
    if (doCancel)
        *cancel = VTRUE;
    else
        *cancel = VFALSE;
        
    // Log the event
    GetSystemTime (&hr, &min, &sec);
    sprintf (buffer, "Workbook BeforeClose event received at %02d::%02d::%02d", 
             hr, min, sec);
    InsertListItem (panelHandle, PANEL_LISTBOX, -1, buffer, 0);
    return S_OK;
}
