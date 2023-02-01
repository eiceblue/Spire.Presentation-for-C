#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"GetShapeGroupAltText.pptx";
	std::wstring outputFile = OutputPath"GetShapeAltText.txt";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load document from disk
	presentation->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);

	//Loop through slides and shapes
	for (int s = 0; s < presentation->GetSlides()->GetCount(); s++)
	{
		ISlide* slide = presentation->GetSlides()->GetItem(s);
		for (int i = 0; i < slide->GetShapes()->GetCount(); i++)
		{
			IShape* shape = slide->GetShapes()->GetItem(i);
			if (dynamic_cast<GroupShape*>(shape) != nullptr)
			{
				//Find the shape group
				GroupShape* groupShape = dynamic_cast<GroupShape*>(shape);
				for (int p = 0; p < groupShape->GetShapes()->GetCount(); p++)
				{
					IShape* gShape = groupShape->GetShapes()->GetItem(p);
					//Append the alternative text in builder
					outFile << gShape->GetAlternativeText() << endl;
				}
			}
		}
	}

	//Write the content in txt file
	outFile.close();
	delete presentation;
}
