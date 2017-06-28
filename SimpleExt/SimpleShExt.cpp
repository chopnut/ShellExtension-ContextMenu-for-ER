// SimpleShExt.cpp : Implementation of CSimpleShExt

#include "stdafx.h"
#include "SimpleShExt.h"



// CSimpleShExt
#define COM_EXCEPTION_GUARD_BEGIN try {

#define COM_EXCEPTION_GUARD_END } catch (const CAtlException & ex){ \
	return static_cast<HRESULT>(ex); \
} catch (const std::bad_alloc &) {\
	return E_OUTOFMEMORY; \
} catch (const std::exception &) {\
	return E_FAIL; \
}
// Helper class but production wise must be in seperate files
class FileEnumFromDataObject{
private:
	//Ban copy as initialize in private
	FileEnumFromDataObject(const FileEnumFromDataObject &);
	FileEnumFromDataObject& operator=(const FileEnumFromDataObject &);

	HDROP m_hDrop; // hdrop stored in idataobject
	STGMEDIUM m_stgm; // "storage" medium (uses a global memory handle)
	UINT m_fileCount; // count of files
public:
	explicit FileEnumFromDataObject(IDataObject *pdtobj)
		:m_hDrop(nullptr) // initialize hdrop handle
		, m_fileCount(0) // initialize number of files
	{
		// Format request hdrop data from the data object
		FORMATETC fmte = {
			CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL
		};

		// Specify the storage medium is global memory handle
		memset(&m_stgm, 0, sizeof(m_stgm));
		m_stgm.tymed = TYMED_HGLOBAL;

		// Retrieve HDRPP data now from source data object
		if (FAILED(pdtobj->GetData(&fmte, &m_stgm))){
			// OutputDebugString(L"[SimpleShell] getdata failed ");

			AtlThrow(E_INVALIDARG);
		}

		// Get a pointer to the file names data
		m_hDrop = reinterpret_cast<HDROP>(::GlobalLock(m_stgm.hGlobal));
		if (!m_hDrop){
			//OutputDebugString(L"[SimpleShell] HDROP is null");
			AtlThrow(E_INVALIDARG);
		}

		// Get number of files
		m_fileCount = ::DragQueryFile(m_hDrop, 0xFFFFFFFF, nullptr, 0);

		CString cs;
		cs.Format(L"[SimpleShell] Files count selected: %d", m_fileCount);
		OutputDebugString(cs);


	}
	// Destructor for resource clean up
	~FileEnumFromDataObject(){
		// if initialization process was successful
		if (m_hDrop){
			// release resources now
			::GlobalUnlock(m_stgm.hGlobal);
			::ReleaseStgMedium(&m_stgm);
		}
	}
	UINT FileCount() const {
		return m_fileCount;
	}
	std::wstring FileAt(UINT index) const
	{
		// Buffer to store filename
		wchar_t filenameBuffer[MAX_PATH];

		// Get current filename
		UINT copiedChars = ::DragQueryFile(
			m_hDrop, 
			index, 
			filenameBuffer,
			_countof(filenameBuffer));

		if ((copiedChars > 0) && (copiedChars < MAX_PATH)){
			return std::wstring(filenameBuffer, copiedChars);
		}
		// On error return an empty string
		return std::wstring();
	}
};
// ** IShellExtInit **
STDMETHODIMP CSimpleShExt::Initialize(
	PCIDLIST_ABSOLUTE pidlFolder,
	IDataObject * pdtobj,
	HKEY hkeyProgID)
{
	// Trap c++ exception and convert them to HRESULT
	COM_EXCEPTION_GUARD_BEGIN


	// Enumerate users selected files
	// extract info from idataobjec

	FileEnumFromDataObject files(pdtobj);
	m_selectedFiles.clear();

	for(UINT i = 0; i < files.FileCount(); ++i){
		std::wstring filename = files.FileAt(i);

		if (!filename.empty()){
			m_selectedFiles.push_back(filename);
		}
	}
	// Trace the initialize code to debugger
	size_t ifilesCount = m_selectedFiles.size();

	// return OK if we have atleast one selected file
	if (ifilesCount!= 0){
		OutputDebugString(L"[SimpleShell] Initialize(S_OK) \r\n");
		return S_OK;
	}
	OutputDebugString(L"[SimpleShell] Initialize(E_INVALIDARG) \r\n");
	return E_INVALIDARG;

	// End exception guard
	COM_EXCEPTION_GUARD_END
}

// ** IContextMenu **
STDMETHODIMP CSimpleShExt::GetCommandString(
	UINT_PTR idCmd,
	UINT uFlags,
	UINT *pwReserved,
	LPSTR pszName,
	UINT cchMax
	){
	COM_EXCEPTION_GUARD_BEGIN;

	// Check that the command ID corresponds to our menu command id
	if (idCmd != m_cmdDisplayFileNames){
		OutputDebugString(_T("[SimpleShell] This is not our command id."));

		return E_INVALIDARG;
	}

	// Check if explorer called us to get the help string
	if (uFlags != GCS_HELPTEXTW){
		// OutputDebugString(_T("[SimpleShell] If not help string."));
		return E_INVALIDARG;

	}

	// Return the help text to the caller

	HRESULT hr = ::StringCchCopy(
		reinterpret_cast<PWSTR>(pszName),
		cchMax,
		m_helpText.c_str());

	return hr;

	COM_EXCEPTION_GUARD_END;
}
STDMETHODIMP CSimpleShExt::QueryContextMenu(
	HMENU hMenu,
	UINT indexMenu,
	UINT idCmdFirst,
	UINT idCmdLast,
	UINT uFlags){
	COM_EXCEPTION_GUARD_BEGIN;
	// If the flags include CMF_DEFAULTONLY, then do nothing
	// as specified in msdn doc, and return control to explorer

	if (uFlags & CMF_DEFAULTONLY){
		return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

	}

	// Modify the menu adding our command
	::InsertMenu(
		hMenu, indexMenu, MF_STRING | MF_BYPOSITION,
		idCmdFirst + m_cmdDisplayFileNames, m_menuText.c_str());

	// Return HRESULT value with severity set to SEVERITY_SUCCESS
	// Set the code value to the offset of the largest command ID
	// that was assigned plus one
	return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL,
		static_cast<USHORT>(m_cmdDisplayFileNames + 1));

	COM_EXCEPTION_GUARD_END;
}
STDMETHODIMP CSimpleShExt::InvokeCommand(
	LPCMINVOKECOMMANDINFO pici)
{
	COM_EXCEPTION_GUARD_BEGIN;

	// We dont use verbs here so exit with an error code
	// if a "verb" was passed
	if (HIWORD(pici->lpVerb) != 0)
		return E_INVALIDARG;
	 
	// Extract the command index from the low-word of lpVerb
	const UINT_PTR cmdIndex = LOWORD(pici->lpVerb);

	// Check that the command index corresponds to our command Id
	if (cmdIndex != m_cmdDisplayFileNames)
		return E_INVALIDARG;


	// ************************
	// Do work below
	if (m_selectedFiles.size() > 0){
		// Step 1: Open stream for writing to the temp file
		std::wstring szFile = CreateFileListIntoFile();

		// Step 2: Read the registry key to find ErnanisRenamer location
		std::wstring szProgram = ReadRegistryKey();

		// Step 3: Create process to run the program with parameter of the path to the temp file.
		if (!szFile.empty() && !szProgram.empty()){
			CreateProcessNow(szProgram, szFile);
		}
	}

	// ************************
	return S_OK;
	COM_EXCEPTION_GUARD_END;

}
void CSimpleShExt::OutputFileNamesToDebugger(void){
	for (const auto& filename : m_selectedFiles){
		std::wstring message = L"Passed file:";
		message += filename;
		message += L"\r\n";

		CString csT;
		csT.Format(L"[SimpleShell] File: %s", message.c_str());
		OutputDebugString(csT);
	}

}
void CSimpleShExt::CreateProcessNow(std::wstring wzProgram, std::wstring wzFile){
	PROCESS_INFORMATION ePI = { 0 };
	STARTUPINFO         rSI = { 0 };
	rSI.cb = sizeof(rSI);
	rSI.dwFlags = STARTF_USESHOWWINDOW;
	rSI.wShowWindow = SW_SHOWNORMAL;  // or SW_HIDE or SW_MINIMIZED

	CString csParam;
	csParam.Format(L"ErnanisRenamer.exe -files %s", wzFile.c_str());

	BOOL fRet = CreateProcess(
		wzProgram.c_str(),  // program name 
		(LPWSTR)csParam.GetString(),     // ...and parameters
		NULL, NULL,  // security stuff (use defaults)
		TRUE,        // inherit handles (not important here)
		0,           // don't need to set priority or other flags
		NULL,        // use default Environment vars
		NULL,        // don't set current directory
		&rSI,        // where we set up the ShowWIndow setting
		&ePI         // gets populated with handle info
		);
	std::wstringstream sw;
	sw << L"[SimpleShell] ProgramPath: " << wzProgram.c_str()<< "\n";
	sw << L"[SimpleShell] ProgramParam: " << csParam.GetString();

	OutputDebugString(sw.str().c_str());
}

std::wstring CSimpleShExt::ReadRegistryKey(void){
	HKEY key; // handle to they key
	wchar_t value[MAX_PATH]; // the container to put registry value
	DWORD value_length = sizeof(value);
	DWORD dwRet;

	dwRet = RegOpenKey(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\ErnanisRenamer"), &key);
	if (dwRet == ERROR_SUCCESS)
	{
		OutputDebugString(L"[SimpleShell] RegOpenKey: Successfully read.");

	}else{
		OutputDebugString(L"[SimpleShell] RegOpenKey: Unsuccessfully read.");
	}

	dwRet = RegQueryValueEx(key, NULL, NULL, NULL, (LPBYTE)&value, &value_length);
	// RegQueryValueEx(
	/*	1.key, handler to the key not a pointer
	2.subkey ,
	3.reserved variable set to null,
	4.type of key,
	5.the value to hold,
	6.value length)
	*/
	if (dwRet == ERROR_SUCCESS)
	{	
		std::wstringstream swsText;
		swsText << L"[SimpleShell] Registry:"<<value;

		OutputDebugString(swsText.str().c_str());
	}
	else{
		char Buffer[MAX_PATH + 1] = { 0 };
		sprintf_s(Buffer, MAX_PATH, "Last Err: %ld", dwRet);

		OutputDebugString(L"[SimpleShell] Something went wrong reading the registry.");
		
	}

	RegCloseKey(key);
	return std::wstring(value);
}
std::wstring CSimpleShExt::CreateFileListIntoFile(void){
	// Create temporary file to store paths
	TCHAR lpTempPath[MAX_PATH];
	TCHAR lpTempFilePath[MAX_PATH];

	DWORD dwRet = GetTempPath(MAX_PATH, lpTempPath);

	// This will populate temfilepath with complete path including the folders
	UINT uiRet = GetTempFileName(lpTempPath, TEXT("rnt"), 0, lpTempFilePath);

	// Clean up the temp file path with double backslash
	CString csTempPath(lpTempFilePath);
	csTempPath.Replace(_T("\\"), _T("\\\\"));


	if (uiRet == 0)
	{
		OutputDebugString(TEXT("[SimpleShell] GetTempFileName failed"));
		return L"";
	}

	CString csT;
	csT.Format(L"[SimpleShell] File: %s", lpTempFilePath);
	OutputDebugString(csT);


	std::wofstream tempStream(lpTempFilePath);

	// Loop through filenames
	for (const auto& filename : m_selectedFiles){
		//std::wstring message = L"Passed file:";
		//message += filename;
		//message += L"\r\n";
		//CString csT;
		//csT.Format(L"[SimpleShell] File: %s", message.c_str());
		//OutputDebugString(csT);

		// Write to the stream
		if (tempStream.is_open()){

			CString csData(filename.c_str());
			csData.Replace(_T('\\'), _T('\\\\'));

			tempStream << csData.GetString() << std::endl;

		}
	}
	tempStream.close(); // delete the temp file by the main program
	return std::wstring(lpTempFilePath);
}


