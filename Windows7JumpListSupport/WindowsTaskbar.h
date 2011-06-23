/*
 * Copyright (C) 2011 by Juan Ramirez (@ichramm) - jramirez@inconcertcc.com
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */
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
