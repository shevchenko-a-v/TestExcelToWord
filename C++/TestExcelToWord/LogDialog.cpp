#include "stdafx.h"
#include "resource.h"
#include "LogDialog.h"



// Window Message Handlers

CLogDialog::CLogDialog(const CString & strLog)
{
	m_strLog = strLog;
	m_strLog.Replace(_T("\n"), _T("\r\n"));// change line endings in order to be correctly shown in Edit control
}

LRESULT CLogDialog::OnCommand(UINT, INT nIdentifier, HWND, BOOL & bHandled)
{
	ATLVERIFY(EndDialog(nIdentifier));
	return 0;
}

LRESULT CLogDialog::OnInitDialog(UINT nMessage, WPARAM wParam, LPARAM lParam, BOOL & bHandled)
{
	ATLVERIFY(CenterWindow());
	m_EditLog.Attach(GetDlgItem(IDC_EDIT_LOG));
	m_EditLog.SetWindowText(m_strLog);
	return 0;
}
