#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SetRowHeightColumnWidth.pptx";
	std::wstring outputFile = OutputPath"SetRowHeightColumnWidth.pptx";

	//Creat a ppt document and load file
	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the table
	ITable* table = nullptr;
	for (int s = 0; s < ppt->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			table = dynamic_cast<ITable*>(shape);

			//Set the height for the rows
			table->GetTableRows()->GetItem(0)->SetHeight(100);
			table->GetTableRows()->GetItem(1)->SetHeight(80);
			table->GetTableRows()->GetItem(2)->SetHeight(60);
			table->GetTableRows()->GetItem(3)->SetHeight(40);
			table->GetTableRows()->GetItem(4)->SetHeight(20);

			//Set the column width
			table->GetColumnsList()->GetItem(0)->SetWidth(60);
			table->GetColumnsList()->GetItem(1)->SetWidth(80);
			table->GetColumnsList()->GetItem(2)->SetWidth(120);
			table->GetColumnsList()->GetItem(3)->SetWidth(140);
			table->GetColumnsList()->GetItem(4)->SetWidth(160);
		}
	}
	//Save the file
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
