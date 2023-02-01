#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Table.pptx";
	std::wstring outputFile = OutputPath"LockAspectRatio.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Verify if it is table
		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			ITable* table = dynamic_cast<ITable*>(shape);
			//Lock aspect ratio
			table->GetShapeLocking()->SetAspectRatioProtection(true);
		}
	}

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
