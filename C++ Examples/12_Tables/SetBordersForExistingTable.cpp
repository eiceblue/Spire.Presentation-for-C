#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"SetBordersForExistingTable.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the table from the first slide of the sample document.
	ISlide* slide = presentation->GetSlides()->GetItem(0);
	ITable* table = dynamic_cast<ITable*>(slide->GetShapes()->GetItem(0));

	//Set the border type as Inside and the border color as blue.
	table->SetTableBorder(TableBorderType::Inside, 1, Color::GetBlue());

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
