// Interop_test.cpp : メイン プロジェクト ファイルです。
// trial to control excel by interop
#include "stdafx.h"
#include <Windows.h>

using namespace System;
using namespace System::Runtime::InteropServices;
using namespace Microsoft::Office::Interop::Excel;
#define Excel  Microsoft::Office::Interop::Excel

int main(array<System::String ^> ^args)
{
	Excel::Application^ xlApp,^xlApp_check;
	Excel::Workbooks^ xlBooks;
	Excel::Workbook^ xlBook;
	Excel::Worksheet^ xlSheet;

	try
	{
		xlApp = static_cast<Excel::Application^>(Marshal::GetActiveObject("Excel::Application"));//already running
	}
	catch (COMException^ ex)
	{
		Console::WriteLine("Staring Excel");

		//run new one
		xlApp = gcnew Excel::ApplicationClass();//run excel
		xlApp->Visible = true;//EnableVisualStyles
	}

	xlBook = xlApp->Workbooks->Add(Type::Missing);
	xlSheet = static_cast<Worksheet^>(xlApp->ActiveSheet);

	//writing on cells
	for (int i = 0; i < 5; i++)
	{
		xlSheet->Cells[i+1,1] = i;
	}
	
	Console::ReadLine();

	xlBook->Close(false, Type::Missing, Type::Missing);//without save
	xlApp->Quit();
	Marshal::ReleaseComObject(xlApp);
	xlApp = nullptr;

	Console::ReadLine();
    return 0;
}
