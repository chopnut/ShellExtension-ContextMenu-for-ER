// SimpleShExt.h : Declaration of the CSimpleShExt

#pragma once
#include "resource.h"       // main symbols
#include "SimpleExt_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// CSimpleShExt


class ATL_NO_VTABLE CSimpleShExt :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSimpleShExt, &CLSID_SimpleShExt>,
	// public ISimpleShExt
	public IShellExtInit,
	public IContextMenu
{
public:
	CSimpleShExt()
		: m_cmdDisplayFileNames(0)
		, m_helpText(L"Rename selected files")
		, m_menuText(L"Rename selected files")
	{
		// COM_INTERFACE_ENTRY(ISimpleShExt)
		OutputDebugString(_T("[SimpleShell] CSimpleShEx Constructor called."));
	};

DECLARE_REGISTRY_RESOURCEID(IDR_SIMPLESHEXT)

DECLARE_NOT_AGGREGATABLE(CSimpleShExt)

BEGIN_COM_MAP(CSimpleShExt)
	COM_INTERFACE_ENTRY(IShellExtInit)
	COM_INTERFACE_ENTRY(IContextMenu)
	// COM_INTERFACE_ENTRY(ISimpleShExt)

END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
		
	}

private:
	// For storing selected files
	std::vector<std::wstring> m_selectedFiles;

	// Shell extension menu command id
	const UINT_PTR m_cmdDisplayFileNames;

	// Help string for our shell extension
	std::wstring m_helpText;

	// Text that will be displayed in the modified context-menu
	std::wstring m_menuText;

	// Helper method: outputs the selected file names to the debugger
	void OutputFileNamesToDebugger(void);
	
	void CreateProcessNow(std::wstring,std::wstring); // To spawn our c# program from registry value
	std::wstring ReadRegistryKey(void); // Read the registry
	std::wstring CreateFileListIntoFile(void); // Create temporary file and return the file


public:
	// IShellExtInit
	STDMETHODIMP Initialize(
		PCIDLIST_ABSOLUTE pidlFolder, 
		IDataObject *pdtobj, 
		HKEY hkeyProgID);

	// IContextMenu
	STDMETHODIMP GetCommandString(
		UINT_PTR idcmd,
		UINT uFlags,
		UINT *pwReserved,
		LPSTR pszName,
		UINT cchMax);
	STDMETHODIMP InvokeCommand(
		LPCMINVOKECOMMANDINFO pici);
	STDMETHODIMP QueryContextMenu(
		HMENU hMenu,
		UINT indexMenu,
		UINT idCmdFirst,
		UINT idCmdLast,
		UINT uFlags);


};

OBJECT_ENTRY_AUTO(__uuidof(SimpleShExt), CSimpleShExt)
