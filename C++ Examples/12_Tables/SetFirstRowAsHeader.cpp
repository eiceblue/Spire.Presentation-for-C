#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"NormalTable.pptx";
	std::wstring outputFile = OutputPath"SetFirstRowAsHeader.pptx";

	ITable* table = nullptr;

	//Load a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			table = dynamic_cast<ITable*>(shape);
		}

	}
	table->SetFirstRow(true);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
