#include "bass.h"
#include "windows.h"
#include <ansi_c.h>
#include <rs232.h>
#include <cvirte.h>		
#include <utility.h>
#include <cvirte.h>		
#include <userint.h>
#include "project.h"
#include "toolbox.h"
#include "exceldem.h"
#include "excel2000.h"

//***** variables for the activex excel *********//
#define APP_WARNING "Warning"
HRESULT ClearObjHandle(CAObjHandle *objHandle);
#define APP_AUTOMATION_ERR "Error:  Microsoft Excel Automation"
#define LAUNCHERR "\
An error occurred trying to launch Excel 2000 through its automation interface.\n\n\
Ensure that Excel is installed and that you can launch it manually. If errors\n\
persist, try to launch Excel manually and use the CONNECT button instead."
static int appVisible = 1;
static ExcelObj_App               ExcelAppHandle = 0;  
static int  UpdateUIRDimming(int panel);
static void ReportAppAutomationError (HRESULT hr);
static ExcelObj_Worksheet         ExcelWorksheetHandle = 0; 
static ExcelObj_Workbooks         ExcelWorkbooksHandle = 0; 
static ExcelObj_Workbook          ExcelWorkbookHandle = 0;  
static ExcelObj_Sheets            ExcelSheetsHandle = 0; 
HRESULT WriteDataToExcel(void);
static VARIANT MyCellRangeV;
#define EXCEL_ARRAY_OF_CELLS "A2:C100"  
static ExcelObj_Range             ExcelRangeHandle = 0;    
static ERRORINFO ErrorInfo;
#define ROWS 98
#define COLUMNS 3
static VARIANT MyVariant;



void CVICALLBACK plotAllAxisToGraph ();

static CmtThreadFunctionID threadid[2]={0};
int CVICALLBACK random(void* rd);
int CVICALLBACK fibonachi(void* rd);

void CVICALLBACK plotXYZForFreeFall ();
void CVICALLBACK serialFunc (int portNumber, int eventMask, void *callbackData);
void CVICALLBACK freeFallDetection ();
static int panelHandle, panelHandle2,panelHandle3;
int com=-1;
int tab_h;
static int PlotGraphsSwitch=0,plotGraphs=0;

static int freeFallswitch=0, plotAllAxis=0;
static double Xarr[1000];
static double Yarr[1000];
static double Zarr[1000];
static int totalTime[1000];
static double allAxisArr[1000];

static double points[3];
HSTREAM freeFallSound;
static double Zaxe,Xaxe,Yaxe;
static double XYZAxis=10; // declare 10  and not 0 so that in the first time freefall will not be detected.
static int arrCount=0;
static int count=1;
static int bytesRead;
static char str[20];
FILE* file;


int main (int argc, char *argv[])
{
	BASS_Init(-1,44100,0,0,NULL); 														//Initializes an output device.	(-1->default device,frequency,not in use,not in use,Class identifier of the object to create -> 0 is a defualt value.
	freeFallSound = BASS_StreamCreateFile(FALSE,"Sound\\freeFallDetected.mp3",0,0,0);	//declare the sound file for freeFall detection
	BASS_ChannelSetAttribute(freeFallSound,BASS_ATTRIB_VOL,1.0);						//Sets the value of a channel's attribute, in this case, the volume.
	if (InitCVIRTE (0, argv, 0) == 0)
		return -1;	/* out of memory */
	if ((panelHandle = LoadPanel (0, "project.uir", PANEL)) < 0)
		return -1;
	if ((panelHandle2 = LoadPanel (0, "project.uir", PANEL2)) < 0)
		return -1;
		if ((panelHandle3 = LoadPanel (0, "project.uir", PANEL_3)) < 0)
		return -1;
	DisplayPanel (panelHandle);
	RunUserInterface ();
	DiscardPanel (panelHandle);
	if (com>=0) CloseCom (com);
	return 0;
}

//// this function connect to the ESP32 via UART with RS232 library.
int CVICALLBACK connect (int panel, int control, int event, void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:
			GetCtrlVal (panelHandle, PANEL_PORT, &com);							//get the com number from the user.
			if (OpenComConfig (com, "", 9600, 0, 8, 1, 512, 512) >=0)			//OpenComConfig Opens a COM port and sets port parameters, the function returns a value that indicate if the action was seccesful. 
			{
				FlushInQ (com);													//Removes all characters from the input queue before recieving data.
				InstallComCallback (com, LWRS_RXFLAG, 0, '\n', serialFunc, 0);	//Allow to install a synchronous callback function for a particular COM port, called when one or more instances of the event character are received and placed in the input queue.
				GetPanelHandleFromTabPage (panelHandle, PANEL_TAB, 0, &tab_h);	//Retrieves the panel handle of a tab page in a tab control.
				SetCtrlAttribute (tab_h, XYZ_Znum, ATTR_DIMMED, 0);				//undim all the controls when connect was pressed after disconnect was pressed.
				SetCtrlAttribute (tab_h, XYZ_Xnum, ATTR_DIMMED, 0);
				SetCtrlAttribute (tab_h, XYZ_Ynum, ATTR_DIMMED, 0);
				SetCtrlAttribute (panelHandle, PANEL_freeFall, ATTR_DIMMED, 0);
				SetCtrlAttribute (panelHandle, PANEL_LAUNCHAPP, ATTR_DIMMED, 0);
				SetCtrlAttribute (panelHandle, PANEL_Connect, ATTR_DIMMED,1);	// only the connect button will be dim.
				SetCtrlAttribute (panelHandle, PANEL_Disconnect, ATTR_DIMMED, 0);
				file = fopen ("output_file.txt", "w");							//open a file to save the data and write headline for each coloum.
				fprintf(file,"		Xaxis			Yaxis			Zaxis			\n");	
			}
			break;
	}
	return 0;
}
//// this function disconnect to the ESP32 via UART with RS232 library.
int CVICALLBACK disconnect (int panel, int control, int event,void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:
			
			if (com>=0)																//if there is a connection.
			{
				CloseCom (com);														//Closes a COM port.
				com = -1;
				GetPanelHandleFromTabPage (panelHandle, PANEL_TAB, 0, &tab_h);		//Retrieves the panel handle of a tab page in a tab control.
				SetCtrlAttribute (tab_h, XYZ_Znum, ATTR_DIMMED, 1);					// dim all the controls except for connect and open excel.
				SetCtrlAttribute (tab_h, XYZ_Xnum, ATTR_DIMMED, 1);
				SetCtrlAttribute (tab_h, XYZ_Ynum, ATTR_DIMMED, 1);
				SetCtrlAttribute (panelHandle, PANEL_freeFall, ATTR_DIMMED, 1);
				SetCtrlAttribute (panelHandle, PANEL_LAUNCHAPP, ATTR_DIMMED, 0);
				SetCtrlAttribute (panelHandle, PANEL_Connect, ATTR_DIMMED, 0);
				SetCtrlAttribute (panelHandle, PANEL_Disconnect, ATTR_DIMMED, 1);
				fclose(file);														//close the file
			}

			break;
	}
	return 0;
}

////this function mannage the incoming data and set the program to work.
void CVICALLBACK serialFunc (int portNumber, int eventMask, void *callbackData)
{
	static char * ptrx,ptry,ptrz;										//to hold the data for each axis.
	count=1;															//counter that indicate if the data is from x,y or z.
	
	while(GetInQLen (com)>0)											//GetInQLen returns the current length of the input queue, if it is bigger than 0, keep getting data.
	{
		totalTime[arrCount]=arrCount;
		bytesRead = ComRdTerm (com, str, 19, '\n');						//Reads from the input queue, returns Number of bytes read from the input queue.get the data to str.
		if (bytesRead<1) 												//no bytes to read, get out of the loop.
			continue;
		str[bytesRead-1]='\0';											//end of the data is a \0.
		GetPanelHandleFromTabPage (panelHandle, PANEL_TAB, 0, &tab_h);	//Retrieves the panel handle of a tab page in a tab control.
		
		if (count == 1)  							// x axis accleration
		{
		 	Xaxe=strtod(str,&ptrx); 						//convert the char data to double.
		 	Xarr[arrCount]= Xaxe;							//get the data into Xarr array.
		 	SetCtrlVal (tab_h, XYZ_Xnum, Xaxe);				//set the gauge to the axis value.
		 	count++;
		}
		else if (count == 2) 	// y axis accleration
		{
			 Yaxe=strtod(str,&ptry);						
			 Yarr[arrCount]= Yaxe;
			 SetCtrlVal (tab_h, XYZ_Ynum, Yaxe); 
			 count++;
		}
		else if (count == 3)	// z axis accleration
		{
			Zaxe=strtod(str,&ptrz);
			Zarr[arrCount]= Zaxe;
			SetCtrlVal (tab_h, XYZ_Znum, Zaxe);

			plotAllAxisToGraph ();						//function call to plot the values in the stripchart.
			plotXYZForFreeFall ();						//function call to plot the values of the free fall calculation.
			fprintf(file,"		%lf		%lf		%lf		\n",Xaxe, Yaxe, Zaxe);	//write the data to the file.
			
			if (arrCount<999)							//get 1000 times the data.
				arrCount++;
			else
				arrCount=0;
			count=1;	
		}
		freeFallDetection ();							//function call for free fall detection.
	}	
}

int CVICALLBACK QuitCallback (int panel, int control, int event,
							  void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:
			QuitUserInterface (0);
			BASS_StreamFree(freeFallSound);				//frees a sample stream's resource		
			break;
	}
	return 0;
}
int CVICALLBACK freeFall (int panel, int control, int event,
						  void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:
			DisplayPanel (panelHandle2);
			break;
	}
	return 0;
}
int CVICALLBACK freeFallSwitch (int panel, int control, int event,void *callbackData, int eventData1, int eventData2)					
{
	switch (event)
	{
		case EVENT_COMMIT:
				GetCtrlVal (panelHandle2, PANEL2_freeFallSwitch, &freeFallswitch);
			break;
	}
	return 0;
}

int CVICALLBACK QuitCallback2 (int panel, int control, int event, void *callbackData, int eventData1, int eventData2)					  
{
	switch (event)
	{
		case EVENT_COMMIT:
		QuitUserInterface (0);
		break;
	}
	return 0;
}


//// function for free fall detection
void CVICALLBACK freeFallDetection ()
		{
		static counter=0;
		XYZAxis = sqrt((Xaxe*Xaxe)+(Yaxe*Yaxe)+(Zaxe*Zaxe)); 		// calculations for free fall.
		allAxisArr[arrCount]= XYZAxis; 								// inserting to array for graph.
		SetCtrlVal (panelHandle2, PANEL2_freeFallAxis, XYZAxis);	// showing on the numeric gauge.
		
		// there are a few samples that reach nearly zero, we need three of them to be sure of a free fall.
		if(counter ==3)
		{
			BASS_ChannelPlay(freeFallSound,TRUE); 					// play a massage "free fall detected"
			counter =0;
		}
		else if(XYZAxis < 2.5)										// if the calculations reached almost zero (we set this to less then 2.5) a freefall ahs accured.
			counter++;
		}
//// function that plots all the axis.
void CVICALLBACK plotAllAxisToGraph ()
{
	GetPanelHandleFromTabPage (panelHandle, PANEL_TAB, 1, &tab_h);						//Retrieves the panel handle of a tab page in a tab control.
	SetCtrlAttribute (tab_h, graphs_StripChart, ATTR_NUM_TRACES, 3);					//sets the strip chart.
	points[0] = Xarr[arrCount];															//value to write in the first trace.
	points[1] = Yarr[arrCount];															//value to write in the second trace.
	points[2] = Zarr[arrCount];															//value to write in the third trace.
	
	int i, numTraces=3;
	for (i=1; i <=numTraces; i++)
		SetTraceAttribute (tab_h, graphs_StripChart, i, ATTR_TRACE_LG_VISIBLE, 1);		//set each trace.
	
	SetTraceAttributeEx (tab_h, graphs_StripChart, 1, ATTR_TRACE_LG_TEXT, "Xaxis");		//set the legend.
	SetTraceAttributeEx (tab_h, graphs_StripChart, 2, ATTR_TRACE_LG_TEXT, "Yaxis");		//set the legend.
	SetTraceAttributeEx (tab_h, graphs_StripChart, 3, ATTR_TRACE_LG_TEXT, "Zaxis");		//set the legend.
	
	 if(PlotGraphsSwitch)					//if the switch to show the graph is on
			{
				GetPanelHandleFromTabPage (panelHandle, PANEL_TAB, 1, &tab_h);							//Retrieves the panel handle of a tab page in a tab control.
				SetAxisScalingMode (tab_h, graphs_StripChart, VAL_LEFT_YAXIS, VAL_MANUAL, -19.6, 19.6);	//set a fixed range of the graph.
				plotGraphs = PlotStripChart (tab_h,graphs_StripChart ,points , 3, 0, 0, VAL_DOUBLE);	//plot the strip chart.													
			}
			else if(plotGraphs)	
			{
				ClearStripChart(tab_h, graphs_StripChart);			//clear the plot if the switch was pressed.
			}
			
}
void CVICALLBACK plotXYZForFreeFall ()
{
		if(freeFallswitch) // if the switch is on
			{
				SetAxisScalingMode (panelHandle2, PANEL2_freeFallGraph, VAL_LEFT_YAXIS, VAL_MANUAL, -19.6, 19.6);	//set a fixed range of the graph.
				SetTraceAttributeEx (panelHandle2, PANEL2_freeFallGraph, 1, ATTR_TRACE_COLOR, VAL_RED);				//set color.
				plotAllAxis =  PlotStripChartPoint (panelHandle2, PANEL2_freeFallGraph, Xarr[arrCount]);			//plot the strip chart.
			}
			else if(plotAllAxis)	
			{
				ClearStripChart (panelHandle2, PANEL2_freeFallGraph); 		//clear the plot if the switch was pressed.
			}	
}
//// this function launchs excel.
int CVICALLBACK LaunchApp (int panel, int control, int event,
						   void *callbackData, int eventData1, int eventData2)
{
	HRESULT error = 0;
    switch (event) {
        case EVENT_COMMIT:
            // Launch App
            SetWaitCursor (1);													// turn on the wait cursor.
            error = Excel_NewApp (NULL, 1, LOCALE_NEUTRAL, 0, &ExcelAppHandle); //create a new _Application object, a negative error code indicates function failure.
            SetWaitCursor (0);													// turn off the wait cursor.
            if (error<0) 
			{
        		MessagePopup (APP_AUTOMATION_ERR, LAUNCHERR);					//massage popup with the error.
				error = 0;
                goto Error;
			}
			
            // Make App Visible
            error = Excel_SetProperty (ExcelAppHandle, NULL, Excel_AppVisible, CAVT_BOOL, appVisible?VTRUE:VFALSE); // set the value of a property of any object in the server.
            if (error<0) 
                goto Error;
    
            UpdateUIRDimming(panelHandle);  //function call to dim some of the buttons.
            MakeApplicationActive ();       // open the excel behind the current panel so the user can open a new file.
            break;
    }
    
Error:    
    if (error < 0) 
	ReportAppAutomationError (error);				//report the error.
    return 0;
}
//// open a new excel file when exel is open.
int CVICALLBACK OpenAppFile (int panel, int control, int event,void *callbackData, int eventData1, int eventData2)
{
    HRESULT error = 0;
    char fileName[MAX_PATHNAME_LEN];
    
    switch (event) {
        case EVENT_COMMIT:
            if (!ExcelWorkbooksHandle)
            {
                // Get Workbooks    
                error = Excel_GetProperty (ExcelAppHandle, NULL, Excel_AppWorkbooks, CAVT_OBJHANDLE, &ExcelWorkbooksHandle); // get the value of a property of any object in the server.
                if (error<0) 
                    goto Error;
    
                // Open existing Workbook
                GetProjectDir (fileName);														//get the directory.
                strcat(fileName, "\\exceldem.xls");
                error = Excel_WorkbooksOpen (ExcelWorkbooksHandle, NULL, fileName, CA_DEFAULT_VAL,	//set the workbook parametrs
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
    
                // Make First Sheet Active. 
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
//// dimming some of the bottuns.
static int UpdateUIRDimming(int panel)
{
    SetCtrlAttribute (panel, PANEL_LAUNCHAPP,    ATTR_DIMMED,  (int)ExcelAppHandle);
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

static void ReportAppAutomationError (HRESULT hr)
{
    char errorBuf[256];
    
    if (hr < 0) {
        CA_GetAutomationErrorString (hr, errorBuf, sizeof (errorBuf));
        MessagePopup (APP_AUTOMATION_ERR, errorBuf);
    }
    return;
}
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
HRESULT WriteDataToExcel(void)
{
    VARIANT *vArray = NULL;
    HRESULT error = 0;
    int i, j;

    SetWaitCursor (1);
    
    // Open new Range for Worksheet
    error = CA_VariantSetCString (&MyCellRangeV, EXCEL_ARRAY_OF_CELLS);   //Converts a C-style string to a BSTR
    error = Excel_WorksheetRange (ExcelWorksheetHandle, NULL, MyCellRangeV, CA_DEFAULT_VAL, &ExcelRangeHandle); //set the range.
    if (error<0) goto Error;

    // Make range Active    
    error = Excel_RangeActivate (ExcelRangeHandle, &ErrorInfo, NULL);
    if (error<0) goto Error;
	
	//insetring data to excel.
        for (j=0;j<ROWS;j++)
        {
            error = Excel_RangeSetItem (ExcelRangeHandle, &ErrorInfo,CA_VariantInt (j+1),CA_VariantInt (1), CA_VariantFloat ( Xarr[j]) );
			error = Excel_RangeSetItem (ExcelRangeHandle, &ErrorInfo,CA_VariantInt (j+1),CA_VariantInt (2), CA_VariantFloat ( Yarr[j]) );
	        error = Excel_RangeSetItem (ExcelRangeHandle, &ErrorInfo,CA_VariantInt (j+1),CA_VariantInt (3), CA_VariantFloat ( Zarr[j]) );
            if (error<0) 
				goto Error;
        }
    

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

int CVICALLBACK plotGraphsSwitch (int panel, int control, int event, void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:
				GetPanelHandleFromTabPage (panelHandle, PANEL_TAB, 1, &tab_h);
				GetCtrlVal (tab_h, graphs_plotGraphsSwitch, &PlotGraphsSwitch);
			break;
	}
	return 0;
}

int CVICALLBACK multiThread (int panel, int control, int event,	 void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:
		DisplayPanel (panelHandle3);
			break;
	}
	return 0;
}
int CVICALLBACK QuitCallback3 (int panel, int control, int event,
							   void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:
			QuitUserInterface (0);
			break;
	}
	return 0;
}
//// multithreading function. this function govern the fibonachi and random.
int CVICALLBACK startDrawing (int panel, int control, int event, void *callbackData, int eventData1, int eventData2)
{
	switch (event)
	{
		case EVENT_COMMIT:
			// start the multithreading.
			CmtScheduleThreadPoolFunction (DEFAULT_THREAD_POOL_HANDLE, fibonachi, NULL, &threadid[0]);
			CmtScheduleThreadPoolFunction (DEFAULT_THREAD_POOL_HANDLE, random, NULL, &threadid[1]);
			CmtWaitForThreadPoolFunctionCompletion ( DEFAULT_THREAD_POOL_HANDLE, threadid[0], OPT_TP_PROCESS_EVENTS_WHILE_WAITING);
			CmtWaitForThreadPoolFunctionCompletion ( DEFAULT_THREAD_POOL_HANDLE, threadid[1], OPT_TP_PROCESS_EVENTS_WHILE_WAITING);
			CmtReleaseThreadPoolFunctionID (DEFAULT_THREAD_POOL_HANDLE, threadid[0]);
			CmtReleaseThreadPoolFunctionID (DEFAULT_THREAD_POOL_HANDLE, threadid[1]);
		
			break;
	}
	return 0;
}

int CVICALLBACK fibonachi(void* rd)
{
	// array that indicate drowing right left up o down
	static int xcord=100,ycord=0;
	int x=0;
	static int num1=0,num2=1,nextNum;
	for(int i=0; i<300; i++) {
		x=!x;
		SetCtrlVal(panelHandle3,PANEL_3_LEDFibonachi  , x);// led blinking
		//fibonachi algoritem
		nextNum=num1+num2;
		num1=num2;
		num2=nextNum;

		SetCtrlAttribute (panelHandle3, PANEL_3_CANVAS, ATTR_PEN_FILL_COLOR, VAL_DK_GREEN);		//set the squere color
		CanvasDrawRect (panelHandle3, PANEL_3_CANVAS, MakeRect (xcord, ycord, num2, num2), VAL_DRAW_FRAME_AND_INTERIOR); //drew the squeres.
		ycord+=num2;
		
		if(num2>144)   // statrt again in num =144.
		{
			CanvasClear (panelHandle3, PANEL_3_CANVAS, VAL_ENTIRE_OBJECT);
			xcord=100;
			ycord=0;
			num1=0;
			num2=1;
		}
		SetCtrlVal (panelHandle3, PANEL_3_NUMERICFibonachi, num1); //showing on the numeric
		Delay(0.07);
	}
	return 0;
}

int CVICALLBACK random(void* rd)
{
	int x=0;
	int num;
	for(int i=0; i<300; i++) 
	{
		x=!x;
		SetCtrlVal(panelHandle3,PANEL_3_LEDRandom  , x);// led blinking
		num = (rand() % 100);
		SetCtrlAttribute (panelHandle3, PANEL_3_CANVAS_2, ATTR_PEN_FILL_COLOR, VAL_DK_RED);
		CanvasDrawOval (panelHandle3, PANEL_3_CANVAS_2, MakeRect (i+rand()%300, i+rand()%300, 20, 20), VAL_DRAW_FRAME_AND_INTERIOR);
		SetCtrlVal (panelHandle3, PANEL_3_NUMERICRandom, num); //showing on the numeric
		Delay(0.05);
	}
	return 0;
}



