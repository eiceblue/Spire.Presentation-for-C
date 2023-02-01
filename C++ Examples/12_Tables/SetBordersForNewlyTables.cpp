#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"SetBordersForNewlyTables.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Set the table width and height for each table cell.
	std::vector<double> tableWidth = { 100, 100, 100, 100, 100 };
	std::vector<double> tableHeight = { 20, 20 };

	//Traverse all the border type of the table.
	for (int i = 1; i <= 2048; i *= 2)
	{
		//Add a table to the presentation slide with the setting width and height
		ITable* itable = presentation->GetSlides()->Append()->GetShapes()->AppendTable(100, 100, tableWidth, tableHeight);

		//Add some text to the table cell.
		itable->GetTableRows()->GetItem(0)->GetItem(0)->GetTextFrame()->SetText(L"Row");
		itable->GetTableRows()->GetItem(1)->GetItem(0)->GetTextFrame()->SetText(L"Column");

		//Set the border type, border width and the border color for the table.
		itable->SetTableBorder(static_cast<TableBorderType>(i), 1.5, Color::GetRed());
	}

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
