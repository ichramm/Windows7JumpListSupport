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
	BSTR dir, file, optsFile;
	m_solution->GetSolutionInfo(&dir, &file, &optsFile);

	m_store.StoreSolution((const TCHAR *)file);
	CWindowsTaskbar::UpdateJumpList(m_store);

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
{
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

