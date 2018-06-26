// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"
#include <tchar.h>

#include <atlbase.h>
#include <atlwin.h>
#include <atlstr.h>

#include "Include\atlapp.h"
#include "Include\atlframe.h"
#include "Include\atlsplit.h"
#include "Include\atlctrls.h"

#import "C:\\Program Files\\Microsoft Office\\root\\vfs\\ProgramFilesCommonX64\\Microsoft Shared\\OFFICE16\\MSO.DLL" \
	rename( "RGB", "MSORGB" ) \
	rename("DocumentProperties", "ExcelDocumentProperties")


#import "C:\\Program Files\\Microsoft Office\\root\\vfs\\ProgramFilesCommonX86\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB"


#import "C:\\Program Files\\Microsoft Office\\root\Office16\\EXCEL.EXE" \
	rename( "DialogBox", "ExcelDialogBox" ) \
	rename( "RGB", "ExcelRGB" ) \
	rename( "CopyFile", "ExcelCopyFile" ) \
	rename( "ReplaceText", "ExcelReplaceText" ) 
