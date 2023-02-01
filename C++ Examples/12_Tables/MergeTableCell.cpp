#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"MergeTableCell.pptx";
	std::wstring outputFile = OutputPath"MergeTableCell.pptx";

	//Create a PPT document and load file
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	ITable* table = nullptr;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			table = dynamic_cast<ITable*>(shape);

			//Merge the second row and third row of the first column
			table->MergeCells(table->GetItem(0, 1), table->GetItem(0, 2), false);

			table->MergeCells(table->GetItem(3, 4), table->GetItem(4, 4), true);
		}
	}
	//Save and launch the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
