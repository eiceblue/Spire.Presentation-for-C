#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"GetShapesByPlaceholder.pptx";
	std::wstring outputFile = OutputPath"GetShapesByPlaceholder.pptx";

	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	Placeholder* placeholder = ppt->GetSlides()->GetItem(1)->GetShapes()->GetItem(0)->GetPlaceholder();
	//Get Shapes by Placeholder
	std::vector<IShape*> shapes = ppt->GetSlides()->GetItem(1)->GetPlaceholderShapes(placeholder);
	std::wstring text = L"";
	//Iterate over all the shapes
	for (int i = 0; i < shapes.size(); i++)
	{
		//If shape is IAutoShape
		if (dynamic_cast<IAutoShape*>(shapes[i]) != nullptr)
		{
			IAutoShape* autoShape = dynamic_cast<IAutoShape*>(shapes[i]);
			if (autoShape->GetTextFrame() != nullptr)
			{
				//text += autoShape->GetTextFrame()->GetText() + "\\r\\n";
				text += autoShape->GetTextFrame()->GetText();
			}
		}
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
