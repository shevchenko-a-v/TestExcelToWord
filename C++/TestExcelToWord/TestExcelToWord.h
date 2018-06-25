#pragma once

////////////////////////////////////////////////////////////
// CMainDialog

class CTestExcelToWordDialog :
	public CDialogImpl<CTestExcelToWordDialog>
{
public:
	enum { IDD = IDD_MAIN };

	BEGIN_MSG_MAP(CMainDialog)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_ID_HANDLER(IDCANCEL, OnCommand)
		COMMAND_ID_HANDLER(IDOK, OnCommand)
		COMMAND_HANDLER(IDC_BTN_BROWSE_SOURCE, BN_CLICKED, OnBnClickedBtnBrowseSource)
		COMMAND_HANDLER(IDC_BTN_BROWSE_OUTPUT, BN_CLICKED, OnBnClickedBtnBrowseOutput)
	END_MSG_MAP()

public:
	// CMainDialog

	// Window Message Handlers
	LRESULT OnInitDialog(UINT nMessage, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		ATLVERIFY(CenterWindow());
		m_EditSourcePath.Attach(GetDlgItem(IDC_EDIT_SOURCE));
		m_EditOutputPath.Attach(GetDlgItem(IDC_EDIT_OUTPUT));
		return 0;
	}
	LRESULT OnCommand(UINT, INT nIdentifier, HWND, BOOL& bHandled)
	{
		ATLVERIFY(EndDialog(nIdentifier));
		return 0;
	}

private:
	LRESULT OnBnClickedBtnBrowseSource(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	
private:
	WTL::CEdit m_EditSourcePath;
	WTL::CEdit m_EditOutputPath;
public:
	LRESULT OnBnClickedBtnBrowseOutput(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
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
