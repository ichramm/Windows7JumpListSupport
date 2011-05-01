#ifndef windowstaskbar_h__
#define windowstaskbar_h__

#pragma once

#include "SolutionStore.h"

#define _IN
#define _OUT

#define MAX_JUMPLIST_ITEMS 30

#ifdef UNICODE
#define IID_IShellLink IID_IShellLinkW
#else
#define IID_IShellLink IID_IShellLinkA
#endif

struct ICustomDestinationList;
struct IShellLink;
struct TaskDescription;

class CWindowsTaskbar
{
public:
	CWindowsTaskbar(void);
	~CWindowsTaskbar(void);

	static void UpdateJumpList(CSolutionStore &store);
	static void ClearJumpList(ICustomDestinationList *list);

private:
	static void CleanupRemovedDestinations(CSolutionStore &store, ICustomDestinationList *list);
	static BOOL SetupShellLinkObject(IShellLink *link, LPCTSTR title, LPCTSTR path, LPCTSTR args, LPCTSTR desc, LPCTSTR iconPath, int iconIndex);
	static IShellLink *CreateTask(TaskDescription *taskDesc, const TCHAR *module);
};

#endif // windowstaskbar_h__
