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
#include "Package.h"
#include "WindowsTaskbar.h"


CWindows7JumpListSupportPackage::CWindows7JumpListSupportPackage()
{
	m_solution = NULL;
	m_pIConnectionPoint = NULL;
}

CWindows7JumpListSupportPackage::~CWindows7JumpListSupportPackage()
{
}

//////////////////////////////////////////////////////////////////////////
// IVsPackage
//////////////////////////////////////////////////////////////////////////

STDMETHODIMP CWindows7JumpListSupportPackage::SetSite(/* [in] */ __RPC__in_opt IServiceProvider *pSP)
{
	if(pSP)
	{
		HRESULT res = pSP->QueryService(SID_SVsSolution, SID_SVsSolution, (void**)&m_solution);
		if(S_OK == res)
		{
			res = m_solution->AdviseSolutionEvents(this, &m_cookie);
		}

		CWindowsTaskbar::UpdateJumpList(m_store);
	}

	return this->IVsPackageImpl::SetSite(pSP);
}

STDMETHODIMP CWindows7JumpListSupportPackage::Close(void)
{
	if(m_solution)
	{
		m_solution->UnadviseSolutionEvents(m_cookie);
		m_solution = NULL;
	}
	return this->IVsPackageImpl::Close();
}

//////////////////////////////////////////////////////////////////////////
// IVsSolutionEvents
//////////////////////////////////////////////////////////////////////////

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnAfterOpenProject(__RPC__in_opt IVsHierarchy *pHierarchy, BOOL fAdded)
{
	return S_OK;
}

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnQueryCloseProject(__RPC__in_opt IVsHierarchy *pHierarchy, BOOL fRemoving, __RPC__inout BOOL *pfCancel)
{
	return S_OK;
}

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnBeforeCloseProject(__RPC__in_opt IVsHierarchy *pHierarchy, BOOL fRemoved)
{
	return S_OK;
}

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnAfterLoadProject(__RPC__in_opt IVsHierarchy *pStubHierarchy, __RPC__in_opt IVsHierarchy *pRealHierarchy)
{
	return S_OK;
}

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnQueryUnloadProject(__RPC__in_opt IVsHierarchy *pRealHierarchy, __RPC__inout BOOL *pfCancel)
{
	return S_OK;
}

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnBeforeUnloadProject(__RPC__in_opt IVsHierarchy *pRealHierarchy, __RPC__in_opt IVsHierarchy *pStubHierarchy)
{
	return S_OK;
}

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnAfterOpenSolution(__RPC__in_opt IUnknown *pUnkReserved, BOOL fNewSolution)
{
	HandleSolutionEvent();
	return S_OK;
}

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnQueryCloseSolution(__RPC__in_opt IUnknown *pUnkReserved, __RPC__inout BOOL *pfCancel)
{
	return S_OK;
}

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnBeforeCloseSolution(__RPC__in_opt IUnknown *pUnkReserved)
{
	return S_OK;
}

STDMETHODIMP 
CWindows7JumpListSupportPackage::OnAfterCloseSolution(__RPC__in_opt IUnknown *pUnkReserved)
{ // handle solution that weren't saved yet (so they don't have a filename before this event is called)
	HandleSolutionEvent();
	return S_OK;
}


// Command handler called when the user selects the "My Command" command.
void CWindows7JumpListSupportPackage::OnMyCommand(CommandHandler* pSender, DWORD flags, VARIANT* pIn, VARIANT* pOut)
{
	// Get the string for the title of the message box from the resource dll.
	CComBSTR bstrTitle;
	VSL_CHECKBOOL_GLE(bstrTitle.LoadStringW(_AtlBaseModule.GetResourceInstance(), IDS_PROJNAME));
	// Get a pointer to the UI Shell service to show the message box.
	CComPtr<IVsUIShell> spUiShell = this->GetVsSiteCache().GetCachedService<IVsUIShell, SID_SVsUIShell>();
	LONG lResult;
	HRESULT hr = spUiShell->ShowMessageBox(
		0,
		CLSID_NULL,
		bstrTitle,
		W2OLE(L"Inside CWindows7JumpListSupportPackage::Exec"),
		NULL,
		0,
		OLEMSGBUTTON_OK,
		OLEMSGDEFBUTTON_FIRST,
		OLEMSGICON_INFO,
		0,
		&lResult);
	VSL_CHECKHRESULT(hr);
}

void CWindows7JumpListSupportPackage::HandleSolutionEvent()
{
	BSTR dir, file, optsFile;
	m_solution->GetSolutionInfo(&dir, &file, &optsFile);

	if(0 != m_lastSolution.CompareNoCase(file))
	{
		m_lastSolution = file;
		m_store.StoreSolution(m_lastSolution);
		CWindowsTaskbar::UpdateJumpList(m_store);
	}
}

