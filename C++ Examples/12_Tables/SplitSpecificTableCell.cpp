#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"SplitSpecificTableCell.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide.
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	//Get the table.
	ITable* table = dynamic_cast<ITable*>(slide->GetShapes()->GetItem(0));

	//Split cell [1, 2) into 3 rows and 2 columns.
	table->GetItem(1, 2)->Split(3, 2);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
