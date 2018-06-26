#pragma once

class CLogDialog :
	public CDialogImpl<CLogDialog>
{
public:
	CLogDialog(const CString &strLog);

	enum { IDD = IDD_LOG };

	BEGIN_MSG_MAP(CMainDialog)
		COMMAND_ID_HANDLER(IDCANCEL, OnCommand)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	END_MSG_MAP()

private:
	// Window Message Handlers
	LRESULT OnCommand(UINT, INT nIdentifier, HWND, BOOL& bHandled);
	LRESULT OnInitDialog(UINT nMessage, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

private:
	WTL::CEdit m_EditLog;
	CString m_strLog;
};