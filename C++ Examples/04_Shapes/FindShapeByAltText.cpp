#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"FindShapeByAltText.pptx";
	std::wstring outputFile = OutputPath"FindShapeByAltText.txt";

	wofstream outFile(outputFile, ios::out);

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	//Find shape in the slide
	IShape* shape = nullptr;

	//Loop through shapes in the slide
	for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
	{
		IShape* shape1 = slide->GetShapes()->GetItem(s);
		//Find the shape whose alternative text is altText

		if (wcscmp(shape1->GetAlternativeText(), L"Shape1") == 0)
		{
			shape = shape1;
		}
	}

	outFile << shape->GetName() << endl;
	delete presentation;
}
