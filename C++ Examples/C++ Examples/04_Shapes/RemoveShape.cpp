#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"FindShapeByAltText.pptx";
	std::wstring outputFile = OutputPath"RemoveShape.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load doucment from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Loop through slides
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		ISlide* slide = presentation->GetSlides()->GetItem(i);
		//Loop through shapes
		for (int j = 0; j < slide->GetShapes()->GetCount(); j++)
		{
			IShape* shape = slide->GetShapes()->GetItem(j);
			//Find the shapes whose alternative text contain "Shape"
			std::wstring temp = shape->GetAlternativeText();
			std::string::size_type pos = temp.find(L"Shape");
			if (pos != string::npos)
			{
				slide->GetShapes()->Remove(shape);
				j--;
			}
		}
	}

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
