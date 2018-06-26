#pragma once


class ExcelToWordConverter
{
public:
	void vTransferExcelToWord(const CString &strSourceFilePath, const CString &strOutputFilePath);

	inline CString strGetLog() { return m_strLog; }

private:
	void vReadFromExcel(const CString &strSourceFile, CString &strLetters, CString &strDigits);
	void vWriteToWord(const CString &strDestinationFile, const CString &strLetters, const CString &strDigits);
	void vWriteLog(const CString& strMessage);

	CString m_strLog;
	const int m_ciMaxDigits = 4;
};

