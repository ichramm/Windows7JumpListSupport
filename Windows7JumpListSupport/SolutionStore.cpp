#include "stdafx.h"
#include "SolutionStore.h"

#include "Package.h"

#define KEY_RECENT_SOLUTIONS		_T("Software\\ichramm\\RecentSolutions")

CSolutionStore::CSolutionStore(void)
{
}

CSolutionStore::~CSolutionStore(void)
{
}

std::vector<CString> CSolutionStore::GetSolutions()
{
	HKEY hKey;
	DWORD nType;
	DWORD nIndex;
	DWORD cchValueName;
	DWORD lpcbData;
	TCHAR valueName[1024+1];
	TCHAR valueData[1024+1];
	std::vector<CString> vect;

	LONG lResult = RegOpenKeyEx(HKEY_CURRENT_USER, KEY_RECENT_SOLUTIONS, 0, KEY_ALL_ACCESS, &hKey);
	if(ERROR_SUCCESS != lResult)
	{ // first time?
		return vect;
	}

	nIndex = 0;
	while (lResult != ERROR_NO_MORE_ITEMS)
	{
		cchValueName = 1024;
		lpcbData = 1024*sizeof(TCHAR);
		lResult = RegEnumValue(hKey, nIndex++, valueName, &cchValueName, 0, &nType, (LPBYTE)valueData, &lpcbData);
		_ASSERTE(ERROR_SUCCESS == lResult || ERROR_NO_MORE_ITEMS == lResult);

		if(lResult == ERROR_SUCCESS && REG_SZ == nType)
		{
			valueData[lpcbData/sizeof(TCHAR)] = _T('0');
			vect.push_back(valueData);
		}
	}
	RegCloseKey(hKey);

	return vect;
}


int CSolutionStore::StoreSolution(const CString &path)
{
	int res = 0;
	HKEY hKey;
	DWORD dwDisposition;
	DWORD nNameIndex;
	DWORD nType;
	TCHAR pvData[1024+1];
	DWORD lpcbData = 1024*sizeof(TCHAR);
	CString fileName = path.Mid(path.ReverseFind(_T('\\'))+1);

	LONG lResult = RegCreateKeyEx(HKEY_CURRENT_USER, KEY_RECENT_SOLUTIONS, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, &hKey, &dwDisposition);
	_ASSERTE(ERROR_SUCCESS == lResult);
	if(ERROR_SUCCESS != lResult) {
		return -1;
	}

	nNameIndex = 2;
	CString keyName = fileName;
	while(ERROR_SUCCESS == RegGetValue(hKey, NULL, (LPCTSTR)keyName, RRF_RT_ANY, &nType, pvData, &lpcbData))
	{
		if(nType == REG_SZ)
		{
			pvData[lpcbData/sizeof(TCHAR)] = 0;
			if (0 == path.CompareNoCase(pvData))
			{
				res = 1;
				break;
			}
		}
		keyName.Format(_T("%s|%d"), fileName, nNameIndex++);
		lpcbData = 1024*sizeof(TCHAR);
	}

	if(!res)
	{
		lResult = RegSetValueEx(hKey, (const TCHAR *)keyName, 0, REG_SZ, (const BYTE*)((LPCTSTR)path), (path.GetLength()+1)*sizeof(TCHAR));
		_ASSERTE(ERROR_SUCCESS == lResult);
		res = ERROR_SUCCESS == lResult ? 0 : -1;
	}

	RegCloseKey(hKey);
	return res;
}

int CSolutionStore::RemoveSolution(const CString &path)
{
	HKEY hKey;
	TCHAR valueName[1024+1];
	TCHAR valueData[1024+1];
	DWORD nType;
	DWORD nEnumIndex;
	DWORD cchValueName = 1024;
	DWORD lpcbData = 1024*sizeof(TCHAR);
	CString keyName;

	LONG lResult = RegOpenKeyEx(HKEY_CURRENT_USER, KEY_RECENT_SOLUTIONS, 0, KEY_ALL_ACCESS, &hKey);
	_ASSERTE(ERROR_SUCCESS == lResult);
	if(ERROR_SUCCESS != lResult) {
		return -1;
	}

	nEnumIndex = 0;
	while (ERROR_SUCCESS == RegEnumValue(hKey, nEnumIndex++, valueName, &cchValueName, 0, &nType, (LPBYTE)valueData, &lpcbData))
	{
		if (REG_SZ == nType)
		{
			valueData[lpcbData/sizeof(TCHAR)] = _T('0');
			if(path.CompareNoCase(valueData) == 0)
			{
				keyName = valueName;
				break;
			}
		}
		cchValueName = 1024;
		lpcbData = 1024*sizeof(TCHAR);
	}

	if(!keyName.IsEmpty())
	{
		RegDeleteKeyValue(hKey, NULL, (LPCTSTR)keyName);
	}

	RegCloseKey(hKey);
	return 0;
}