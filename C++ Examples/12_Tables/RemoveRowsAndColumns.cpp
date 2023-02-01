#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"RemoveRowsAndColumns.pptx";
	std::wstring outputFile = OutputPath"RemoveRowsAndColumns.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the table in PPT document
	ITable* table = nullptr;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			table = dynamic_cast<ITable*>(shape);

			//Remove the second column
			table->GetColumnsList()->RemoveAt(1, false);

			//Remove the second row
			table->GetTableRows()->RemoveAt(1, false);
		}
	}
	//Save and launch the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
