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

#define  PANEL                            1
#define  PANEL_LAUNCH                     2       /* callback function: Launch */
#define  PANEL_TABLE                      3       /* callback function: TableCB */
#define  PANEL_COPYTABLE                  4       /* callback function: CopyTable */
#define  PANEL_CALCULATE                  5       /* callback function: Calculate */
#define  PANEL_GRAPH                      6       /* callback function: Graph */
#define  PANEL_RING                       7       /* callback function: ChartSelect */
#define  PANEL_QUIT                       8       /* callback function: Quit */


     /* Menu Bars, Menus, and Menu Items: */

          /* (no menu bars in the resource file) */


     /* Callback Prototypes: */ 

int  CVICALLBACK Calculate(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK ChartSelect(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK CopyTable(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK Graph(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK Launch(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK Quit(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK TableCB(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);


#ifdef __cplusplus
    }
#endif
