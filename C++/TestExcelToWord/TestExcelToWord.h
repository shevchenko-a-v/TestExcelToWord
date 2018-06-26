#pragma once
#include "ExcelToWordConverter.h"
////////////////////////////////////////////////////////////
// CMainDialog

class CTestExcelToWordDialog :
	public CDialogImpl<CTestExcelToWordDialog>
{
public:
	enum { IDD = IDD_MAIN };

	BEGIN_MSG_MAP(CMainDialog)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_HANDLER(IDC_BTN_BROWSE_SOURCE, BN_CLICKED, OnBnClickedBtnBrowseSource)
		COMMAND_HANDLER(IDC_BTN_BROWSE_OUTPUT, BN_CLICKED, OnBnClickedBtnBrowseOutput)
		COMMAND_HANDLER(IDC_BTN_TRANSFER, BN_CLICKED, OnBnClickedBtnTransfer)
	END_MSG_MAP()

private:
	// Window Message Handlers
	LRESULT OnInitDialog(UINT nMessage, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	LRESULT OnBnClickedBtnBrowseSource(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedBtnBrowseOutput(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedBtnTransfer(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	
private:
	WTL::CEdit m_EditSourcePath;
	WTL::CEdit m_EditOutputPath;

	ExcelToWordConverter m_converter;
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
