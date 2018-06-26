#include "stdafx.h"
#include "ExcelToWordConverter.h"
#include <chrono>
#include <stdexcept>
#include <experimental\filesystem>
#include <atlpath.h>
#include <regex>


using namespace std;
typedef std::chrono::system_clock Clock;


void ExcelToWordConverter::vTransferExcelToWord(const CString & strSourceFilePath, const CString & strOutputFilePath)
{
	m_strLog.Empty();
	try
	{
		vWriteLog(_T("Started transfer from[") + strSourceFilePath + _T("] to [") + strOutputFilePath + _T("]"));
		CPath pathSrc(strSourceFilePath);
		CPath pathDst(strOutputFilePath);
		if (!pathSrc.FileExists())
			throw invalid_argument("Source Excel file does not exist.");
		if (pathDst.FileExists())
		{
				vWriteLog(_T("Removing destination file"));
				if (DeleteFile(strOutputFilePath))
					vWriteLog(_T("Destination file is successfully removed."));
				else
					vWriteLog(_T("Destination file cannot be removed. We will try to overrite it."));
		}

		CString strLetters, strDigits;
		vReadFromExcel(strSourceFilePath, strLetters, strDigits);
		vWriteToWord(strOutputFilePath, strLetters, strDigits);
	}
	catch (_com_error& e)
	{
		vWriteLog(e.ErrorMessage());
		throw;
	}
	catch (exception& e)
	{
		vWriteLog(CString(e.what()));
		throw;
	}	
}

void ExcelToWordConverter::vReadFromExcel(const CString & strSourceFile, CString & strLetters, CString & strDigits)
{
	try
	{
		strLetters = _T("");
		strDigits = _T("");
		vWriteLog(_T("Started reading from source file."));

		Excel::_ApplicationPtr pApplication;
		if (FAILED(pApplication.CreateInstance(_T("Excel.Application"))))
			throw runtime_error("Excel could not be started. Check that you have Microsoft Office installed.");
		vWriteLog(_T("Excel is started."));

		vWriteLog(_T("Opening workbook..."));
		Excel::_WorkbookPtr pWorkbook = pApplication->Workbooks->Open(_com_util::ConvertStringToBSTR(CStringA(strSourceFile)), 0, true);
		vWriteLog(_T("Workbook is successfully opened."));

		vWriteLog(_T("Opening worksheet..."));
		Excel::_WorksheetPtr pWorksheet = pWorkbook->Sheets->Item[1];
		vWriteLog(_T("Worksheet is successfully opened."));
		auto rowsNumber = pWorksheet->UsedRange->Rows->Count;

		wregex regLetters(_T("^[[:alpha:]]+$"), regex_constants::extended);
		wregex regDigits(_T("^\\d+$"));
		vWriteLog(_T("Obtaining values from the first column started."));
		for (long i = 0; i < rowsNumber; ++i)
		{
			_variant_t  vItem = pWorksheet->UsedRange->Item[1+i][1];
			CString strValue((LPTSTR)_bstr_t(vItem));
			if (regex_match((LPCTSTR)strValue, regLetters))
			{
				if (!strLetters.IsEmpty())
					strLetters += _T(" ");
				strLetters += strValue;
			}
			else if (regex_match((LPCTSTR)strValue, regDigits))
			{
				if (!strDigits.IsEmpty())
					strDigits += _T("-");
				CString strDigitBlock(strValue);
				if (strDigitBlock.GetLength() > m_ciMaxDigits)
					strDigitBlock.Delete(m_ciMaxDigits, strDigitBlock.GetLength() - m_ciMaxDigits);
				strDigits += strDigitBlock;
			}
		}
		vWriteLog(_T("Obtaining values from the first column completed."));
		pWorkbook->Close(VARIANT_FALSE); 
		pApplication->Quit();
		vWriteLog(_T("Excel is closed."));
	}
	catch(...)
	{
		vWriteLog(_T("Error occured during reading from source file."));
		throw;
	}
	vWriteLog(_T("Reading from source file completed successfully."));
}

void ExcelToWordConverter::vWriteToWord(const CString & strDestinationFile, const CString & strLetters, const CString & strDigits)
{
	try
	{
		vWriteLog(_T("Started writing to destination file."));

		Word::_ApplicationPtr pApplication;
		if (FAILED(pApplication.CreateInstance(_T("Word.Application"))))
			throw runtime_error("Word could not be started. Check that you have Microsoft Office installed."); vWriteLog(_T("Excel is started."));

		vWriteLog(_T("Adding new document..."));
		Word::_DocumentPtr pDoc = pApplication->Documents->Add();
		vWriteLog(_T("New document is successfully added."));

		pDoc->Content->Font->Size = 12;
		CString strContent(strLetters + _T("\n") + strDigits);
		pDoc->Content->Text = _bstr_t(strContent);

		vWriteLog(_T("Text added."));
		
		CComVariant path(strDestinationFile);
		pDoc->SaveAs2(&path);
		pDoc->Close();
		vWriteLog(_T("Word file has been saved to disk."));
		pApplication->Quit();
		vWriteLog(_T("Word is closed."));

	}
	catch(...)
	{
		vWriteLog(_T("Error occured during writing to destination file."));
		throw;
	}
	vWriteLog(_T("Writing to destination file completed successfully."));
}

void ExcelToWordConverter::vWriteLog(const CString & strMessage)
{
	auto now = Clock::now();
	auto seconds = std::chrono::time_point_cast<std::chrono::seconds>(now);
	auto fraction = now - seconds;
	time_t cnow = Clock::to_time_t(now); 
	auto milliseconds = std::chrono::duration_cast<std::chrono::milliseconds>(fraction);
	
	std::tm tm;
	localtime_s(&tm, &cnow);


	TCHAR buffer[10];
	_tcsftime(buffer, 10, _T("%H:%M:%S"), &tm);
	CString strLogEntry;
	strLogEntry.Format(_T("%s.%03d\t"), buffer, milliseconds.count());
	strLogEntry += strMessage + _T("\n");
	m_strLog += strLogEntry;
}
