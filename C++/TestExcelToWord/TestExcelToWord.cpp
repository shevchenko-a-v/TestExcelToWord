#include "stdafx.h"
#include "resource.h"
#include "TestExcelToWord.h"
#include "ExcelToWordConverter.h"


CTestExcelToWordDialogModule _AtlModule;

////////////////////////////////////////////////////////////
// Main

extern "C" int WINAPI _tWinMain(HINSTANCE /*hInstance*/, HINSTANCE /*hPrevInstance*/, LPTSTR /*lpCmdLine*/, int nShowCmd)
{
	return _AtlModule.WinMain(nShowCmd);
}




// Window Message Handlers

LRESULT CTestExcelToWordDialog::OnInitDialog(UINT nMessage, WPARAM wParam, LPARAM lParam, BOOL & bHandled)
{
	ATLVERIFY(CenterWindow());
	m_EditSourcePath.Attach(GetDlgItem(IDC_EDIT_SOURCE));
	m_EditOutputPath.Attach(GetDlgItem(IDC_EDIT_OUTPUT));
	return 0;
}

LRESULT CTestExcelToWordDialog::OnBnClickedBtnBrowseSource(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	TCHAR buffer[2048];
	buffer[0] = '\0';
	OPENFILENAME of; 
	ZeroMemory(&of, sizeof(of));

	of.lStructSize = sizeof(OPENFILENAME);
	of.lpstrFilter = _T("Excel files(*.xlsx)\0*.xlsx\0\0");
	of.nFileOffset = 1;
	of.nMaxFile = 2048;
	of.lpstrFile = buffer;
	of.lpstrTitle = _T("Select source Excel file");
	of.Flags = OFN_DONTADDTORECENT | OFN_FILEMUSTEXIST | OFN_NONETWORKBUTTON;
	if (GetOpenFileName(&of))
	{
		ATLASSERT(m_EditSourcePath.IsWindow());
		m_EditSourcePath.SetWindowText(buffer);
	}
	return 0;
}


LRESULT CTestExcelToWordDialog::OnBnClickedBtnBrowseOutput(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	TCHAR buffer[2048];
	buffer[0] = '\0';
	OPENFILENAME of;
	ZeroMemory(&of, sizeof(of));

	of.lStructSize = sizeof(OPENFILENAME);
	of.lpstrFilter = _T("Word files(*.docx)\0*.docx\0\0");
	of.nFileOffset = 1;
	of.lpstrDefExt = _T("docx");
	of.nMaxFile = 2048;
	of.lpstrFile = buffer;
	of.lpstrTitle = _T("Specify output Word file");
	of.Flags = OFN_DONTADDTORECENT  | OFN_NONETWORKBUTTON | OFN_OVERWRITEPROMPT;

	if (GetSaveFileName(&of))
	{
		ATLASSERT(m_EditOutputPath.IsWindow());
		CString strFileName(buffer);
		if (of.Flags & OFN_EXTENSIONDIFFERENT)
		{
			int iExtensionPos = strFileName.ReverseFind(_T('.'));
			if (iExtensionPos >= 0)
			{
				strFileName.Delete(iExtensionPos, strFileName.GetLength() - iExtensionPos);
				strFileName += ".docx";
			}
		}
		m_EditOutputPath.SetWindowText(strFileName);
	}
	return 0;
}


LRESULT CTestExcelToWordDialog::OnBnClickedBtnTransfer(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	ATLASSERT(m_EditSourcePath.IsWindow());
	ATLASSERT(m_EditOutputPath.IsWindow());

	CString strSource, strDestination;
	m_EditSourcePath.GetWindowText(strSource);
	m_EditOutputPath.GetWindowText(strDestination);

	m_converter.vTransferExcelToWord(strSource, strDestination);

	return 0;
}
