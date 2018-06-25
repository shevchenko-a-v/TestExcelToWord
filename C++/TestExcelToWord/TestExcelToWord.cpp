#include "stdafx.h"
#include "resource.h"

////////////////////////////////////////////////////////////
// CMainDialog

class CTestExcelToWordDialog:
	public CDialogImpl<CTestExcelToWordDialog>
{
public:
	enum { IDD = IDD_MAIN };

	BEGIN_MSG_MAP(CMainDialog)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(IDCANCEL, OnCommand)
		COMMAND_ID_HANDLER(IDOK, OnCommand)
	END_MSG_MAP()

public:
	// CMainDialog

	// Window Message Handlers
	LRESULT OnInitDialog(UINT nMessage, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		ATLVERIFY(CenterWindow());
		return 0;
	}
	LRESULT OnCommand(UINT, INT nIdentifier, HWND, BOOL& bHandled)
	{
		ATLVERIFY(EndDialog(nIdentifier));
		return 0;
	}
};

////////////////////////////////////////////////////////////
// CTestExcelToWordDialogModule

class CTestExcelToWordDialogModule :
	public CAtlExeModuleT<CTestExcelToWordDialogModule>
{
public:
	// CTestExcelToWordDialogModule
	HRESULT PreMessageLoop(INT nShowCommand)
	{
		_ATLTRY
		{
			ATLENSURE_SUCCEEDED(__super::PreMessageLoop(nShowCommand));
		}
			_ATLCATCH(Exception)
		{
			return Exception;
		}
		return S_OK;
	}
	VOID RunMessageLoop()
	{
		CTestExcelToWordDialog Dialog;
		Dialog.DoModal();
	}
};

CTestExcelToWordDialogModule _AtlModule;

////////////////////////////////////////////////////////////
// Main

extern "C" int WINAPI _tWinMain(HINSTANCE /*hInstance*/, HINSTANCE /*hPrevInstance*/, LPTSTR /*lpCmdLine*/, int nShowCmd)
{
	return _AtlModule.WinMain(nShowCmd);
}

