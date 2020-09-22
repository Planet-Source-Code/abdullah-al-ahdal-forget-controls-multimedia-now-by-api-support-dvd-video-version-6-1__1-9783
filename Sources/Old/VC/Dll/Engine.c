// Engine.cpp : Defines the entry point for the DLL application.
//

#include "stdio.h"
#include "windows.h"
#include "mmsystem.h"


HANDLE ghMod;
#define BUFSIZE 80
DWORD glo_from;
DWORD glo_to;
HWND glo_hWnd ;


VOID APIENTRY CloseMPEG();
VOID APIENTRY PauseMPEG();
VOID APIENTRY ResumeMPEG();
VOID APIENTRY StopMPEG();
CHAR APIENTRY MoveMPEG(DWORD to);
CHAR APIENTRY CALLBACK  OpenMPEG(HWND hwnd ,CHAR FileName[255],CHAR typeAviOrMPEG[6]);
CHAR APIENTRY PlayMPEG(DWORD from,DWORD to);
CHAR APIENTRY PutMPEG(INT left,INT top,INT Width,INT Height);
CHAR APIENTRY GetDefaultDevice(char typeDevice[20]);
LONG APIENTRY GetCurrentMPEGPos();
LONG APIENTRY GetTotalframes();
LONG APIENTRY GetTotalTimeByms();
LONG APIENTRY GetFramesPerSecond();
DWORD APIENTRY GetPercent();
LONG APIENTRY GetStatusMPEG();
BOOL APIENTRY AreMPEGAtEnd();
void APIENTRY SetAutoRepeat(int autorep);
void TimerFunction(HWND hwnd,unsigned int a,unsigned int b,unsigned long c);
void APIENTRY SetDefaultDevice(char typeDevice[20],char drvDefaultDevice[20]);


BOOL APIENTRY DllMain( HANDLE hDLL, 
                       DWORD  dwReason, 
                       LPVOID lpReserved
					 )
{
  switch (dwReason)
  {
    case DLL_PROCESS_ATTACH:
    {
		char buf[BUFSIZE+1];

      //
      // DLL is attaching to the address space of the current process.
      //

      ghMod = hDLL;
      GetModuleFileName (NULL, (LPTSTR) buf, BUFSIZE);
      break;
    }

    case DLL_THREAD_ATTACH:

      //
      // A new thread is being created in the current process.
      //

      break;

    case DLL_THREAD_DETACH:

      //
      // A thread is exiting cleanly.
      //

  	  SetAutoRepeat (-1);
	  CloseMPEG();

      break;

    case DLL_PROCESS_DETACH:

      //
      // The calling process is detaching the DLL from its address space.
      //
	  SetAutoRepeat (-1);
	  CloseMPEG();

      break;
  }

return TRUE;
}


CHAR WINAPI CALLBACK  OpenMPEG(HWND hwnd ,CHAR FileName[255],CHAR typeAviOrMPEG[6])
{
char cmdToDo[255];
DWORD dwReturn;        // mciSendString return value
CHAR ret[8];

char shortpath[255];

GetShortPathName(FileName , shortpath , 255);


glo_hWnd=hwnd ; //store hwnd in global buffer

sprintf(cmdToDo,"open %s Type %s Alias mpeg parent %d Style 1073741824",shortpath ,typeAviOrMPEG,hwnd);
dwReturn=mciSendString(cmdToDo,NULL, 0, NULL);


if (! dwReturn == 0)//not success
	mciGetErrorString (dwReturn,ret,128);

else
	strcpy(ret,"Success");


return ret;
}

CHAR WINAPI CALLBACK  PlayMPEG(DWORD from,DWORD to)
{
	DWORD dwReturn;        // mciSendString return value
	char cmdToDo[255];
	CHAR ret[128];

	if (from==NULL && to ==NULL)
	{

		glo_from=1;						//store the value in global buffer
		glo_to=GetTotalframes();		//store the value in global buffer
		goto complete;
		
	}
	else if (! from == NULL && ! to ==NULL)
	{
	
		glo_from=from;	//store the value in global buffer
		glo_to=to;		//store the value in global buffer
		goto complete;

	}
	else if ( to == NULL && ! from == NULL)
	{

		glo_from=from;					//store the value in global buffer
		glo_to=GetTotalframes ();		//store the value in global buffer
		goto complete;
	}
		
	else if ( !to == NULL &&  from == NULL)
	{

		glo_from=1;	//store the value in global buffer
		glo_to=to;		//store the value in global buffer
		goto complete;
	}

complete:
	sprintf(cmdToDo,"play mpeg from %d to %d",glo_from,glo_to);

	dwReturn=mciSendString(cmdToDo,NULL, 0, NULL);

if (! dwReturn == 0)//not success
	mciGetErrorString (dwReturn,ret,128);

else
	strcpy(ret,"Success");
return ret;
}

CHAR APIENTRY PutMPEG(INT left,INT top,INT Width,INT Height)
{

	DWORD dwReturn;        // mciSendString return value
	char cmdToDo[255];
	CHAR ret[128];
	if (left == NULL)
		left=0;
	if (top	 == NULL)
		top=0;

	if ( Width == NULL || Height ==NULL)
	{
		RECT rect;
		GetWindowRect (glo_hWnd,(LPRECT)&rect);
		Width = rect.right - rect.left ;
		Height= rect.bottom - rect.top ;
	}



	sprintf(cmdToDo ,"put mpeg window at %d %d %d %d",left,top,Width ,Height);

	dwReturn=mciSendString(cmdToDo,NULL, 0, NULL);


if (! dwReturn == 0)//not success
	mciGetErrorString (dwReturn,ret,128);

else
	strcpy(ret,"Success");
return ret;
}

CHAR APIENTRY GetDefaultDevice(char typeDevice[20])
{

char tmp[255];
CHAR device[255];
char path[255];

GetWindowsDirectory(tmp, 255);
sprintf(path ,"%s\system.ini",tmp);

GetPrivateProfileString("MCI", typeDevice, "None", device , 255, path);
return device;

}

void APIENTRY SetDefaultDevice(char typeDevice[20],char drvDefaultDevice[20])
{

char tmp[255];
char path[255];

GetWindowsDirectory(tmp, 255);
sprintf(path ,"%s\\system.ini",tmp);

WritePrivateProfileString("MCI", typeDevice,drvDefaultDevice ,path);


}


VOID APIENTRY CloseMPEG()
{
mciSendString("Close mpeg",NULL, 0, NULL);
}

VOID APIENTRY StopMPEG()
{
mciSendString("Stop mpeg",NULL, 0, NULL);
}

VOID APIENTRY PauseMPEG()
{
mciSendString("Pause mpeg",NULL, 0, NULL);
}

VOID APIENTRY ResumeMPEG()
{
	DWORD dwReturn;        // mciSendString return value
	CHAR ret[128];
	dwReturn=mciSendString("Resume mpeg",NULL, 0, NULL);

if (! dwReturn == 0)//not success
	mciGetErrorString (dwReturn,ret,128);

else
	strcpy(ret,"Success");


return ret;
}

CHAR APIENTRY MoveMPEG(DWORD to)
{
	DWORD dwReturn;        // mciSendString return value
	char cmdToDo[255];
	CHAR ret[128];
	sprintf(cmdToDo,"seek mpeg to %d",to);
	dwReturn=mciSendString(cmdToDo,NULL, 0, NULL);
	mciSendString("Play mpeg",NULL, 0, NULL);

if (! dwReturn == 0)//not success
	mciGetErrorString (dwReturn,ret,128);

else
	strcpy(ret,"Success");
return ret;

}

LONG APIENTRY GetStatusMPEG()
{
	DWORD dwReturn;        // mciSendString return value
	char status[128];
	CHAR ret[128];
	SHORT result;

	dwReturn=mciSendString("status mpeg mode", status, 128, NULL);
	result=strcmp(status,"playing");
	if (result==0)
			return 3;
	result=strcmp(status,"paused");
	if (result==0)
			return 2;
	result=strcmp(status,"stopped");
	if (result==0)
			return 1;		

	return -1;
}

LONG APIENTRY GetTotalTimeByms()
{
	DWORD dwReturn;        // mciSendString return value
	char timeMS[128];
	int result;
	

	mciSendString("set mpeg time format ms", timeMS, 128, NULL);
	dwReturn=mciSendString("status mpeg length", timeMS, 128, NULL);

	if (! dwReturn == 0)//if not success
	return -1;

	//Success
	return atol(timeMS);
}

LONG APIENTRY GetTotalframes()
{
	DWORD dwReturn;        // mciSendString return value
	char total[128];
	int result;

	mciSendString("set mpeg time format frames", total, 128, NULL);
	dwReturn=mciSendString("status mpeg length", total, 128, NULL);

	if (! dwReturn == 0)//if not success
	return -1;

	//Success
	return atol(total);	
}

LONG APIENTRY GetCurrentMPEGPos()
{
	DWORD dwReturn;        // mciSendString return value
	char pos[128];
	int result;
	
	dwReturn=mciSendString("status mpeg position", pos, 128, NULL);

	if (! dwReturn == 0)//if not success
	return -1;

	//Success
	return atol(pos);
}

DWORD APIENTRY GetPercent()
{

	DWORD totalframes;
	DWORD currframe;
	currframe=GetCurrentMPEGPos () ;
	totalframes=GetTotalframes ();
	return currframe * 100 / totalframes;
}

LONG APIENTRY GetFramesPerSecond()
{
DWORD totalframes;
DWORD totalTime;
totalTime	 = GetTotalTimeByms();
totalframes  = GetTotalframes();

if (totalframes == -1 || totalTime  == -1) 
    return -1;
    

return totalframes / (totalTime / 1000);


}

BOOL APIENTRY AreMPEGAtEnd()
{
		DWORD currpos;
		currpos=GetCurrentMPEGPos ();
		if (glo_to == currpos || (glo_to -1) < currpos)
			return TRUE;
		else
			return FALSE;
}

void APIENTRY SetAutoRepeat(int autorep)
{
if (autorep==1)
	SetTimer (glo_hWnd ,50,100,(TIMERPROC)TimerFunction);

else
	KillTimer(glo_hWnd,50);
}


void TimerFunction(HWND hwnd,unsigned int a,unsigned int b,unsigned long c)
{
			DWORD currpos;
			currpos=GetCurrentMPEGPos ();

			if (glo_to == currpos || (glo_to -1) < currpos)
			{

				PlayMPEG (glo_from,glo_to);

			}

}



