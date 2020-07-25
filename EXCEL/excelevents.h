/**************************************************************************/
/* LabWindows/CVI User Interface Resource (UIR) Include File              */
/*                                                                        */
/* WARNING: Do not add to, delete from, or otherwise modify the contents  */
/*          of this include file.                                         */
/**************************************************************************/

#include <userint.h>

#ifdef __cplusplus
    extern "C" {
#endif

     /* Panels and Controls: */

#define  PANEL                           1
#define  PANEL_LAUNCHBUTTON              2       /* callback function: LaunchExcel */
#define  PANEL_CLOSEBUTTON               3       /* callback function: ReleaseExcel */
#define  PANEL_HANDLECLOSEBOOKBUTTON     4       /* callback function: HandleCloseBook */
#define  PANEL_DETACHCLOSEBOOKBUTON      5       /* callback function: DetachFromCloseBook */
#define  PANEL_CANCELCLOSEBOOKSWITCH     6
#define  PANEL_LISTBOX                   7
#define  PANEL_QUITBUTTON                8       /* callback function: Quit */
#define  PANEL_TEXTMSG                   9


     /* Menu Bars, Menus, and Menu Items: */

          /* (no menu bars in the resource file) */


     /* Callback Prototypes: */ 

int  CVICALLBACK DetachFromCloseBook(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK HandleCloseBook(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK LaunchExcel(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK Quit(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK ReleaseExcel(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);


#ifdef __cplusplus
    }
#endif
