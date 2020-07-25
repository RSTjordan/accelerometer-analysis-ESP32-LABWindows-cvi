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
#define  PANEL_LAUNCHAPP                 2       /* callback function: LaunchApp */
#define  PANEL_CONNECTAPP                3       /* callback function: ConnectApp */
#define  PANEL_SHUTDOWNAPP               4       /* callback function: ShutdownApp */
#define  PANEL_VISIBILITY                5       /* callback function: ChangeVisibility */
#define  PANEL_OPENFILE                  6       /* callback function: OpenAppFile */
#define  PANEL_SAVEFILE                  7       /* callback function: SaveAppFile */
#define  PANEL_PRINTFILE                 8       /* callback function: PrintAppFile */
#define  PANEL_CLOSEFILE                 9       /* callback function: CloseAppFile */
#define  PANEL_WRITEDATA                 10      /* callback function: WriteData */
#define  PANEL_READDATA                  11      /* callback function: ReadData */
#define  PANEL_MAKECHART                 12      /* callback function: MakeChart */
#define  PANEL_RUNMACRO                  13      /* callback function: RunMacro */
#define  PANEL_QUIT                      14      /* callback function: Quit */
#define  PANEL_DECORATION_2              15
#define  PANEL_TEXTMSG                   16
#define  PANEL_TEXTMSG_2                 17
#define  PANEL_DECORATION                18
#define  PANEL_TEXTMSG_3                 19
#define  PANEL_DECORATION_3              20


     /* Menu Bars, Menus, and Menu Items: */

          /* (no menu bars in the resource file) */


     /* Callback Prototypes: */ 

int  CVICALLBACK ChangeVisibility(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK CloseAppFile(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK ConnectApp(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK LaunchApp(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK MakeChart(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK OpenAppFile(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK PrintAppFile(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK Quit(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK ReadData(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK RunMacro(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK SaveAppFile(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK ShutdownApp(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK WriteData(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);


#ifdef __cplusplus
    }
#endif
