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
#ifndef package_h__
#define package_h__
#pragma once

#include <atlstr.h>
#include <VSLCommandTarget.h>
#include "resource.h"       // main symbols
#include "Guids.h"
#include "..\Windows7JumpListSupportUI\Resource.h"
#include "..\Windows7JumpListSupportUI\CommandIds.h"
#include "SolutionStore.h"
using namespace VSL;

/*
Package Load Key (PLK)	
ATHAM8E0EJI3CIMPI1JAJ3QQMCERJMZRKRIKJ3CJMAQ0ZKKKJ8H2JAC9C1CHH3ETEAJPDEQQHCQHQZAMJ2AHEZA8A3MJDKR8M0J3MED3DHP9P9APK2P0CRPQKDPID3D2
Company Name	ichramm
Package Name	Windows 7 Jump List Support
Package GUID	{a82752e2-f1e0-48ff-adeb-6e7ca1626c21}
PLK Version	1.0
Min. Visual Studio Version	Visual Studio 2008
Minimum Product Edition	Standard
*/

class ATL_NO_VTABLE CWindows7JumpListSupportPackage : 
	// CComObjectRootEx and CComCoClass are used to implement a non-thread safe COM object, and 
	// a partial implementation for IUknown (the COM map below provides the rest).
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CWindows7JumpListSupportPackage, &CLSID_Windows7JumpListSupport>,
	// Provides the implementation for IVsPackage to make this COM object into a VS Package.
	public IVsPackageImpl<CWindows7JumpListSupportPackage, &CLSID_Windows7JumpListSupport>,
	// Show up on the splash screen and the Help About dialog.
	public IVsInstalledProductImpl<IDS_OFFICIALNAME, IDS_VERSION, IDS_PRODUCTDETAILS, IDI_PACKAGE_ICON>,
	public IVsSolutionEvents,
	public IOleCommandTargetImpl<CWindows7JumpListSupportPackage>,
	// Provides consumers of this object with the ability to determine which interfaces support extended error information.
	public ATL::ISupportErrorInfoImpl<&__uuidof(IVsPackage)>
{

BEGIN_COM_MAP(CWindows7JumpListSupportPackage)
	COM_INTERFACE_ENTRY(IVsPackage)
	COM_INTERFACE_ENTRY(IVsInstalledProduct)
	COM_INTERFACE_ENTRY(IOleCommandTarget)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IVsSolutionEvents)
END_COM_MAP()

VSL_BEGIN_REGISTRY_MAP(IDR_REGISTRYSCRIPT)
	VSL_REGISTRY_MAP_GUID_ENTRY_EX(CLSID_Windows7JumpListSupport, CLSID_Package)
	VSL_REGISTRY_MAP_NUMBER_ENTRY(IDS_PACKAGE_LOAD_KEY)
	VSL_REGISTRY_RESOURCEID_ENTRY(IDS_OFFICIALNAME)
	VSL_REGISTRY_RESOURCEID_ENTRY(IDS_PRODUCTDETAILS)
VSL_END_REGISTRY_MAP()

VSL_DECLARE_NOT_COPYABLE(CWindows7JumpListSupportPackage)

public:
	CWindows7JumpListSupportPackage();
	~CWindows7JumpListSupportPackage();

// overrides from IVsPackageImpl
public:
	STDMETHOD(SetSite)(__RPC__in_opt IServiceProvider *pSP);
	STDMETHOD(Close)(void);

// implementation of IVsSolutionEvents
public:
	STDMETHOD(OnAfterOpenProject)   (IVsHierarchy *pHierarchy, BOOL fAdded);
	STDMETHOD(OnQueryCloseProject)  (IVsHierarchy *pHierarchy, BOOL fRemoving, BOOL *pfCancel);
	STDMETHOD(OnBeforeCloseProject) (IVsHierarchy *pHierarchy, BOOL fRemoved);
	STDMETHOD(OnAfterLoadProject)   (IVsHierarchy *pStubHierarchy, IVsHierarchy *pRealHierarchy);
	STDMETHOD(OnQueryUnloadProject) (IVsHierarchy *pRealHierarchy, BOOL *pfCancel);
	STDMETHOD(OnBeforeUnloadProject)(IVsHierarchy *pRealHierarchy, IVsHierarchy *pStubHierarchy);
	STDMETHOD(OnAfterOpenSolution)  (IUnknown *pUnkReserved, BOOL fNewSolution); // the one
	STDMETHOD(OnQueryCloseSolution) (IUnknown *pUnkReserved, BOOL *pfCancel);
	STDMETHOD(OnBeforeCloseSolution)(IUnknown *pUnkReserved);
	STDMETHOD(OnAfterCloseSolution) (IUnknown *pUnkReserved);

	// This function provides the error information if it is not possible to load
	// the UI dll. It is for this reason that the resource IDS_E_BADINSTALL must
	// be defined inside this dll's resources.
	static const LoadUILibrary::ExtendedErrorInfo& GetLoadUILibraryErrorInfo()
	{
		static LoadUILibrary::ExtendedErrorInfo errorInfo(IDS_E_BADINSTALL);
		return errorInfo;
	}

// Definition of the commands handled by this package
VSL_BEGIN_COMMAND_MAP()
    VSL_COMMAND_MAP_ENTRY(CLSID_Windows7JumpListSupportCmdSet, CmdIdJumpListSupport, NULL, CommandHandler::ExecHandler(&OnMyCommand))
VSL_END_VSCOMMAND_MAP()

// Command handler called when the user selects the "My Command" command.
void OnMyCommand(CommandHandler* pSender, DWORD flags, VARIANT* pIn, VARIANT* pOut);


private:

	// called when a solution is opened or closed
	void HandleSolutionEvent();

	VSCOOKIE              m_cookie;
	IVsSolution          *m_solution;
	IConnectionPoint     *m_pIConnectionPoint;
	CSolutionStore        m_store;
	CString               m_lastSolution; // filename of the current opened solution
};

// This exposes CWindows7JumpListSupportPackage for instantiation via DllGetClassObject; however, an instance
// can not be created by CoCreateInstance, as CWindows7JumpListSupportPackage is specifically registered with
// VS, not the the system in general.
OBJECT_ENTRY_AUTO(CLSID_Windows7JumpListSupport, CWindows7JumpListSupportPackage)


#endif // package_h__
