#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"FindShapeByAltText.pptx";
	std::wstring outputFile = OutputPath"HideShape.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Loop through slides
	for (int l = 0; l < presentation->GetSlides()->GetCount(); l++)
	{
		ISlide* slide = presentation->GetSlides()->GetItem(l);
		//Loop through shapes in the slide
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			IShape* shape = slide->GetShapes()->GetItem(s);
			//Find the shape whose alternative text is Shape1
			if (wcscmp(shape->GetAlternativeText(), L"Shape1") == 0)
			{
				//Hide the shape
				shape->SetIsHidden(true);
			}
		}
	}
	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
