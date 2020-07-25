#include "ExcelReport.h"
#include "excel2000.h"
#include <cvirte.h>		
#include <userint.h>
#include "ExcelReportDemo.h"

static int panelHandle;
static CAObjHandle applicationHandle = 0;
static CAObjHandle workbookHandle = 0;
static CAObjHandle worksheetHandle = 0;
static CAObjHandle chartHandle = 0;
static int running = 0;
static int copytableDone = 0;
static int PlotType = ExRConst_GalleryArea;

#define APP_AUTOMATION_ERR "Error:  Microsoft Excel Automation"
#define LAUNCHERR "\
An error occurred trying to launch Excel through its automation interface.\n\n\
Ensure that Excel is installed and that you can launch it manually."

int main (int argc, char *argv[])
{
	if (InitCVIRTE (0, argv, 0) == 0)
		return -1;	/* out of memory */
	if ((panelHandle = LoadPanel (0, "ExcelReportDemo.uir", PANEL)) < 0)
		return -1;
	DisplayPanel (panelHandle);
	RunUserInterface ();
	DiscardPanel (panelHandle);
	return 0;
}

//****************************************************************************************************
//		Launch
//****************************************************************************************************

int CVICALLBACK Launch (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	HRESULT error = 0;

	switch (event)
	{
		case EVENT_COMMIT:
			error = ExcelRpt_ApplicationNew(1, &applicationHandle);
			if (error<0) 
			{
        		MessagePopup (APP_AUTOMATION_ERR, LAUNCHERR);
                goto Error;
			}
			ExcelRpt_WorkbookNew(applicationHandle, &workbookHandle);
			ExcelRpt_WorksheetNew(workbookHandle, -1, &worksheetHandle);
			SetCtrlAttribute (panelHandle, PANEL_LAUNCH, ATTR_DIMMED, 1);
			SetCtrlAttribute (panelHandle, PANEL_TABLE, ATTR_DIMMED, 0);
			SetCtrlAttribute (panelHandle, PANEL_COPYTABLE, ATTR_DIMMED, 0);
			break;
	}

Error:
	return 0;
}

//****************************************************************************************************
//		CopyTable
//****************************************************************************************************

int CVICALLBACK CopyTable (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	switch (event)
		{
		case EVENT_COMMIT:
			ExcelRpt_WriteDataFromTableControl (worksheetHandle, "B2:C7",
														 panelHandle, PANEL_TABLE);
			
			ExcelRpt_RangeBorder (worksheetHandle, "B2:C7", 
										   ExRConst_Continuous, 255, ExRConst_Thin,
										   ExRConst_InsideHorizontal | ExRConst_InsideVertical | 
										   ExRConst_EdgeBottom |ExRConst_EdgeLeft|ExRConst_EdgeRight|
										   ExRConst_EdgeTop);
	
			SetCtrlAttribute (panelHandle, PANEL_COPYTABLE, ATTR_DIMMED, 1);
			SetCtrlAttribute (panelHandle, PANEL_CALCULATE, ATTR_DIMMED, 0);
			SetCtrlAttribute (panelHandle, PANEL_GRAPH, ATTR_DIMMED, 0);
			SetCtrlAttribute (panelHandle, PANEL_RING, ATTR_DIMMED, 1);
	
			copytableDone = 1;
			break;
		}
	return 0;
}

//****************************************************************************************************
//		Calculate
//****************************************************************************************************

int CVICALLBACK Calculate (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	switch (event)
		{
		case EVENT_COMMIT:
			ExcelRpt_SetCellRangeAttribute (worksheetHandle, "B8",
													 ER_CR_ATTR_FORMULA, "=AVERAGE(B2:B7)");
		
			ExcelRpt_SetCellRangeAttribute (worksheetHandle, "C8",
													 ER_CR_ATTR_FORMULA, "=AVERAGE(C2:C7)");
		
			ExcelRpt_SetCellRangeAttribute (worksheetHandle, "A1", ER_CR_ATTR_COLUMN_WIDTH, 10.0);
			
			ExcelRpt_SetCellValue (worksheetHandle, "A8",
											ExRConst_dataString, "AVERAGE:");
																	
			SetCtrlAttribute (panelHandle, PANEL_CALCULATE, ATTR_DIMMED, 1);
			
			break;
		case EVENT_RIGHT_CLICK:
			break;
		}
	return 0;
}

//****************************************************************************************************
//		Graph
//****************************************************************************************************

int CVICALLBACK Graph (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	char *categoryAxisTitle = NULL;
	char *valueAxisTitle = NULL;
    
    switch (event)
		{
		case EVENT_COMMIT:
			
			ExcelRpt_SetCellRangeAttribute (worksheetHandle, "A2:A10",
											ER_CR_ATTR_COLUMN_WIDTH, 5.0);
			
			ExcelRpt_ChartAddtoWorksheet (worksheetHandle, 2.0 * 72, 7.0,
												   4.0*72, 4.0*72, &chartHandle);
				
			// These plot types do not support axis titles.
			if (PlotType != ExRConst_GalleryPie && 
				PlotType != ExRConst_Gallery3DPie && 
				PlotType != ExRConst_GalleryRadar &&
				PlotType != ExRConst_GalleryDoughnut)
				{
					categoryAxisTitle = "Data";
					valueAxisTitle = "Value";
				}
            ExcelRpt_ChartWizard (chartHandle, worksheetHandle, "B2:C7", PlotType, 
                                            0, 0, 0, 0, 1, "Value Data", categoryAxisTitle, valueAxisTitle, NULL);
			
			ExcelRpt_SetCellRangeAttribute (worksheetHandle, "A1", ER_CR_ATTR_COLUMN_WIDTH, 10.0);
			
			ExcelRpt_SetChartAttribute (chartHandle, ER_CH_ATTR_PLOTAREA_COLOR,
										0xffffff);
			
			SetCtrlAttribute (panelHandle, PANEL_GRAPH, ATTR_DIMMED, 1);
			SetCtrlAttribute (panelHandle, PANEL_RING, ATTR_DIMMED, 0);

			running = 1;
			
			break;

		}
	return 0;
}

//****************************************************************************************************
//		Quit
//****************************************************************************************************

int CVICALLBACK Quit (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	switch (event)
		{
		case EVENT_COMMIT:
			if (worksheetHandle)
				CA_DiscardObjHandle(worksheetHandle);
			if (chartHandle)
				CA_DiscardObjHandle(chartHandle);
			if (workbookHandle)
			{
				ExcelRpt_WorkbookClose(workbookHandle, 0);
				CA_DiscardObjHandle(workbookHandle);
			}
			if (applicationHandle)
			{
				ExcelRpt_ApplicationQuit(applicationHandle);
				CA_DiscardObjHandle(applicationHandle);
			}
			QuitUserInterface (0);
			break;
		}
	return 0;
}

//****************************************************************************************************
//		TableCB
//****************************************************************************************************

int CVICALLBACK TableCB (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	switch (event)
		{
		case EVENT_COMMIT:
			
			if(copytableDone)
				ExcelRpt_WriteDataFromTableControl (worksheetHandle, "B2:C7", panelHandle, PANEL_TABLE);
			
			if(running)
			{
				char *categoryAxisTitle = NULL;
				char *valueAxisTitle = NULL;
				
				// These plot types do not support axis titles.
				if (PlotType != ExRConst_GalleryPie && 
					PlotType != ExRConst_Gallery3DPie && 
					PlotType != ExRConst_GalleryRadar &&
					PlotType != ExRConst_GalleryDoughnut)
				{
					categoryAxisTitle = "Data";
					valueAxisTitle = "Value";
				}
	            ExcelRpt_ChartWizard (chartHandle, worksheetHandle, "B2:C7", PlotType, 
	                                            0, 0, 0, 0, 1, "Value Data", categoryAxisTitle, valueAxisTitle, NULL);
			}
			break;
		}
	return 0;
}

//****************************************************************************************************
//		ChartSelect
//****************************************************************************************************

int CVICALLBACK ChartSelect (int panel, int control, int event,
		void *callbackData, int eventData1, int eventData2)
{
	switch (event)
		{
		case EVENT_COMMIT:
			GetCtrlVal (panel, PANEL_RING, &PlotType);
			if (running) {
				char *categoryAxisTitle = NULL;
				char *valueAxisTitle = NULL;
				
				// These plot types do not support axis titles.
				if (PlotType != ExRConst_GalleryPie && 
					PlotType != ExRConst_Gallery3DPie && 
					PlotType != ExRConst_GalleryRadar &&
					PlotType != ExRConst_GalleryDoughnut)
				{
					categoryAxisTitle = "Data";
					valueAxisTitle = "Value";
				}
	            ExcelRpt_ChartWizard (chartHandle, worksheetHandle, "B2:C7", PlotType, 
	                                            0, 0, 0, 0, 1, "Value Data", categoryAxisTitle, valueAxisTitle, NULL);
			}
		break;
		}
	return 0;
}
