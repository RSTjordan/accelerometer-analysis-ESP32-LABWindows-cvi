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
#define  PANEL_LAUNCHAPP                  2       /* control type: command, callback function: LaunchApp */
#define  PANEL_OPENFILE                   3       /* control type: command, callback function: OpenAppFile */
#define  PANEL_WRITEDATA                  4       /* control type: command, callback function: WriteData */
#define  PANEL_DECORATION_3               5       /* control type: deco, callback function: (none) */
#define  PANEL_DECORATION_2               6       /* control type: deco, callback function: (none) */
#define  PANEL_freeFall                   7       /* control type: command, callback function: freeFall */
#define  PANEL_multiThread                8       /* control type: command, callback function: multiThread */
#define  PANEL_Disconnect                 9       /* control type: command, callback function: disconnect */
#define  PANEL_Connect                    10      /* control type: command, callback function: connect */
#define  PANEL_PORT                       11      /* control type: numeric, callback function: (none) */
#define  PANEL_TAB                        12      /* control type: tab, callback function: (none) */
#define  PANEL_QUITBUTTON                 13      /* control type: command, callback function: QuitCallback */
#define  PANEL_Excel                      14      /* control type: textMsg, callback function: (none) */

#define  PANEL2                           2
#define  PANEL2_freeFallAxis              2       /* control type: scale, callback function: (none) */
#define  PANEL2_freeFallSwitch            3       /* control type: binary, callback function: freeFallSwitch */
#define  PANEL2_QUITBUTTON2               4       /* control type: command, callback function: QuitCallback2 */
#define  PANEL2_freeFallGraph             5       /* control type: strip, callback function: (none) */

#define  PANEL_3                          3
#define  PANEL_3_QUITBUTTON3              2       /* control type: command, callback function: QuitCallback3 */
#define  PANEL_3_CANVAS_2                 3       /* control type: canvas, callback function: (none) */
#define  PANEL_3_CANVAS                   4       /* control type: canvas, callback function: (none) */
#define  PANEL_3_startDrawing             5       /* control type: command, callback function: startDrawing */
#define  PANEL_3_LEDFibonachi             6       /* control type: LED, callback function: (none) */
#define  PANEL_3_LEDRandom                7       /* control type: LED, callback function: (none) */
#define  PANEL_3_NUMERICRandom            8       /* control type: numeric, callback function: (none) */
#define  PANEL_3_NUMERICFibonachi         9       /* control type: numeric, callback function: (none) */

     /* tab page panel controls */
#define  graphs_plotGraphsSwitch          2       /* control type: binary, callback function: plotGraphsSwitch */
#define  graphs_StripChart                3       /* control type: strip, callback function: (none) */

     /* tab page panel controls */
#define  XYZ_Ynum                         2       /* control type: scale, callback function: (none) */
#define  XYZ_Xnum                         3       /* control type: scale, callback function: (none) */
#define  XYZ_Znum                         4       /* control type: scale, callback function: (none) */


     /* Control Arrays: */

          /* (no control arrays in the resource file) */


     /* Menu Bars, Menus, and Menu Items: */

          /* (no menu bars in the resource file) */


     /* Callback Prototypes: */

int  CVICALLBACK connect(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK disconnect(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK freeFall(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK freeFallSwitch(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK LaunchApp(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK multiThread(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK OpenAppFile(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK plotGraphsSwitch(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK QuitCallback(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK QuitCallback2(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK QuitCallback3(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK startDrawing(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);
int  CVICALLBACK WriteData(int panel, int control, int event, void *callbackData, int eventData1, int eventData2);


#ifdef __cplusplus
    }
#endif