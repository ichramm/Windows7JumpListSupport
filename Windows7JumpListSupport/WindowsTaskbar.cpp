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
#include "stdafx.h"
#include "ShObjIdl.h"
#include "WindowsTaskbar.h"
#include <algorithm>


struct TaskDescription
{
	LPCTSTR title;
	LPCTSTR path; // NULL use module filename
	LPCTSTR args;
	LPCTSTR desc;
	LPCTSTR iconPath;  // NULL use module filename
	int     iconIndex;
};

static TaskDescription predefined_tasks[] = {
	{ _T("Create Project"), NULL, _T("/Command File.NewProject"), _T("Create a new Visual Studio Solution"), NULL, -1200 },
	{ _T("Open Project"), NULL,  _T("/Command File.OpenProject"), _T("Open Project/Solution"), NULL, -1200 }
};

//Using as extern fails to find symbol
//    extern const PROPERTYKEY PKEY_Title;
// PropertyKey(new Guid("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"), 2);
static const GUID GUID_PKEY_Title = {0xF29F85E0, 0x4FF9, 0x1068, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9};
static const PROPERTYKEY PKEY_Title = { GUID_PKEY_Title , 2};


static __forceinline bool CHECK_SUCCEEDED(HRESULT hr, const TCHAR *errMsg)
{
	bool ret = SUCCEEDED(hr);
	if(!ret)
	{
		OutputDebugString(errMsg);
	}
	return ret;
}

CWindowsTaskbar::CWindowsTaskbar(void)
{
}

CWindowsTaskbar::~CWindowsTaskbar(void)
{
}

/************************************************************************/
/************************************************************************/
BOOL CWindowsTaskbar::SetupShellLinkObject(IShellLink *link, LPCTSTR title, LPCTSTR path, LPCTSTR args, LPCTSTR desc, LPCTSTR iconPath, int iconIndex)
{
	if(!path || !title) {
		return FALSE;
	}
	if(args != NULL) {
		link->SetArguments((const TCHAR *)args);
	}
	if(desc != NULL) {
		link->SetDescription((const TCHAR *)desc);
	}
	if(iconPath != NULL) {
		link->SetIconLocation((const TCHAR *)iconPath, iconIndex);
	}
	IPropertyStore *propertyStore;
	link->SetPath((const TCHAR *)path);
	HRESULT hr = link->QueryInterface(IID_IPropertyStore, (LPVOID *)&propertyStore);
	if (CHECK_SUCCEEDED(hr, _T("SetupShellLinkObject - Failed to QueryInterface IPropertyStore on link")))
	{
		PROPVARIANT propTitle;
		PropVariantInit(&propTitle);
#ifdef UNICODE
		propTitle.vt = VT_LPWSTR;
		propTitle.pwszVal = StrDup(title);
#else
		propTitle.vt = VT_LPSTR;
		propTitle.pszVal = StrDup(title);
#endif
		propertyStore->SetValue(PKEY_Title, propTitle);
		propertyStore->Commit();
#ifdef UNICODE
		LocalFree(propTitle.pwszVal);
#else
		LocalFree(propTitle.pszVal);
#endif
		propertyStore->Release();
	}
	return TRUE;
}

/************************************************************************/
/***********          don't forget to ->Release()             ***********/
IShellLink *CWindowsTaskbar::CreateTask(TaskDescription *taskDesc, LPCTSTR module = NULL)
{
	IShellLink *obj;
	HRESULT hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID *)&obj);
	if (CHECK_SUCCEEDED(hr, _T("CreateTask - Failed to create ShellLink object")))
	{
		if (TRUE == SetupShellLinkObject(obj, taskDesc->title, taskDesc->path ? taskDesc->path : module, taskDesc->args,
			taskDesc->desc, taskDesc->iconPath ? taskDesc->iconPath : module, taskDesc->iconIndex))
		{
			return obj;
		}
		// failed to setup, release
		obj->Release();
	}
	return NULL;
}

/************************************************************************/
/************************************************************************/
void CWindowsTaskbar::UpdateJumpList(CSolutionStore &store)
{
	HRESULT hr;
	
	// Our main jump list instance
	ICustomDestinationList *jumpList;

	hr = CoCreateInstance(CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, IID_ICustomDestinationList, (LPVOID *)&jumpList);
	if(!CHECK_SUCCEEDED(hr, _T("UpdateJumpList - Failed to create customDestinationList")))
	{
		return;
	}

	// if the user has removed items we should update the registry
	CleanupRemovedDestinations(store, jumpList);

	// make sure the jump list is not broken...
	// the jump list gets corrupt when an item is removed
	// Alt: erasing these files seems to fix the problem too:
	//  %AppData%\Microsoft\Windows\Recent\AutomaticDestinations\ed3d6a1919d189ef.automaticDestinations-ms
	//  %AppData%\Microsoft\Windows\Recent\CustomDestinations\12213a23a507245e.customDestinations-ms
	// EDIT: Comment this 'cause it seems that if the items are removed correctly there is no problem
	//ClearJumpList(customDestinationList);

	// the maximum count of items the user has set
	UINT maxUserItems;
	// Items removed by user from jump list. We clean up the jump list on each update so this will alway be empty
	IObjectArray *removedItems;
	hr = jumpList->BeginList(&maxUserItems, IID_IObjectArray, (LPVOID *)&removedItems);
	if (CHECK_SUCCEEDED(hr, _T("UpdateJumpList - Failed to do BeginList")))
	{
		std::vector<IUnknown*>  vectObjects;

		TCHAR buffer[MAX_PATH];
		GetModuleFileName(NULL, buffer, _countof(buffer) - 1);
		CString moduleFilename = buffer;

		IObjectCollection *customCategoryCollection;
		hr = CoCreateInstance(CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC_SERVER, IID_IObjectCollection, (LPVOID *)&customCategoryCollection);
		if (CHECK_SUCCEEDED(hr, _T("UpdateJumpList - Failed to create customCategoryCollection")))
		{
			vectObjects.push_back(customCategoryCollection);

			//GetEnvironmentVariable(_T("CommonProgramFiles"), buffer, _countof(buffer) - 1);
			//PathAppend(buffer, _T("Microsoft Shared\\MSEnv\\VSLauncher.exe"));
			//CString vsLauncherPath = (GetFileAttributes(buffer) != INVALID_FILE_ATTRIBUTES) ? buffer : moduleFilename;

			std::vector<CString> solutions = store.GetSolutions();

			int counter = 0;
			for ( ; counter < (int)solutions.size() && counter < std::min(MAX_JUMPLIST_ITEMS, (int)maxUserItems); counter++)
			{
				const CString &filePath = solutions[counter];

				if (GetFileAttributes((LPCTSTR)filePath) == INVALID_FILE_ATTRIBUTES)
				{ // if the solution file does not exist forget about it
					store.RemoveSolution(filePath);
					solutions.erase(solutions.begin() + counter-- );
					continue;
				}
				
				CString fileName = filePath.Mid(filePath.ReverseFind(_T('\\'))+1);
				CString safeFilePath = filePath;
				if (safeFilePath.Find(_T(' '), 1) > 0)
				{
					safeFilePath.Insert(0, _T("\""));
					safeFilePath.Append(_T("\""));
				}

				TaskDescription task = { (LPCTSTR)fileName, NULL, (LPCTSTR)safeFilePath, (LPCTSTR)filePath, NULL, 3 };
				IShellLink *current = CreateTask(&task, (LPCTSTR)moduleFilename);
				if(current)
				{
					customCategoryCollection->AddObject(current);
					vectObjects.push_back(current);
				}
			}
			
			// append the category and commit the list
			IObjectArray *jumpListItemsAsArray;
			hr = customCategoryCollection->QueryInterface(IID_IObjectArray, (LPVOID *)&jumpListItemsAsArray);
			if (CHECK_SUCCEEDED(hr, _T("UpdateJumpList - Failed to QueryInterface IObjectArray on customCategoryCollection")))
			{ // append the custom category to the jump list
				jumpList->AppendCategory(L"Recent Solutions", jumpListItemsAsArray);
				jumpListItemsAsArray->Release();
			}
		}

		IObjectCollection *taskCollection = NULL;
		hr = CoCreateInstance(CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC_SERVER, IID_IObjectCollection, (LPVOID *)&taskCollection);
		if (CHECK_SUCCEEDED(hr, _T("UpdateJumpList - Failed to create taskCollection")))
		{
			vectObjects.push_back(taskCollection);
			for (int j = 0; j < _countof(predefined_tasks); j++)
			{
				IShellLink *currentTask = CreateTask(&predefined_tasks[j], (LPCTSTR)moduleFilename);
				if(currentTask)
				{
					taskCollection->AddObject(currentTask);
					vectObjects.push_back(currentTask);
				}
			}

			IObjectArray *taskArray;
			hr = taskCollection->QueryInterface(IID_IObjectArray, (LPVOID *)&taskArray);
			if (CHECK_SUCCEEDED(hr, _T("UpdateJumpList - Failed to QueryInterface IObjectArray on taskCollection")))
			{
				jumpList->AddUserTasks(taskArray);
				taskArray->Release();
			}
		}

		// commit the changes
		jumpList->CommitList();

		// release all COM objects (shell links + collections)
		for ( ; !vectObjects.empty(); vectObjects.pop_back()) {
			vectObjects.back()->Release();
		}

		removedItems->Release();
	}

	jumpList->Release();
}

/************************************************************************/
/************************************************************************/
void CWindowsTaskbar::ClearJumpList(ICustomDestinationList *jumpList)
{
	if(jumpList)
	{
		jumpList->AddRef();
	}
	else
	{
		HRESULT hr = CoCreateInstance(CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, IID_ICustomDestinationList, (LPVOID *)&jumpList);
		if(!CHECK_SUCCEEDED(hr, _T("ClearJumpList - Failed to create customDestinationList")))
		{
			return;
		}
	}

	/* We don't use DeleteList, because that requires an explicitly set AppID. */
	UINT num_items;
	IObjectArray *pRemoved;
	if (SUCCEEDED(jumpList->BeginList(&num_items, IID_IObjectArray, (LPVOID *)&pRemoved)))
	{
		jumpList->CommitList();
		pRemoved->Release();
	}

	jumpList->Release();
}

/************************************************************************/
/************************************************************************/
void CWindowsTaskbar::CleanupRemovedDestinations(CSolutionStore &store, ICustomDestinationList *list)
{
	IObjectArray *removedItems;
	HRESULT hr = list->GetRemovedDestinations(IID_IObjectArray, (LPVOID *)&removedItems);
	if(CHECK_SUCCEEDED(hr, _T("CleanupRemovedDestinations - Failed to GetRemovedDestinations on list")))
	{
		UINT remCount;
		HRESULT hr = removedItems->GetCount(&remCount);
		if(SUCCEEDED(hr))
		{
			CString msg;
			for (UINT i = 0; i < remCount; i++)
			{
				IShellLink *removedLink = NULL;
				hr = removedItems->GetAt(i, IID_IShellLink, (LPVOID *)&removedLink);
				if(CHECK_SUCCEEDED(hr, _T("CleanupRemovedDestinations - Failed to GetAt on removedItems")))
				{
					TCHAR buffer[MAX_PATH];
					ZeroMemory(buffer, sizeof(buffer));
					hr = removedLink->GetArguments(buffer, _countof(buffer));
					if(SUCCEEDED(hr)) {
						store.RemoveSolution(buffer);
					}
					removedLink->Release();
				}
			}
		}
		removedItems->Release();
	}
}
