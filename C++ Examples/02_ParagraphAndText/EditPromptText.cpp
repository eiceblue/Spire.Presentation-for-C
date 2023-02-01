#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"HasPromptText.pptx";
	std::wstring outputFile = OutputPath"EditPromptText.pptx";

	//Load a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	// Iterate through the slide
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++) {
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (shape->GetPlaceholder() != nullptr && dynamic_cast<IAutoShape*>(shape) != nullptr)
		{
			std::wstring text = L"";
			// Set the text of the title
			if (shape->GetPlaceholder()->GetType() == PlaceholderType::CenteredTitle)
			{
				text = L"custom title create by Spire";
			}
			// Set text of the subtitle.
			else if (shape->GetPlaceholder()->GetType() == PlaceholderType::Subtitle)
			{
				text = L"custom subtitle create by Spire";
			}

			(dynamic_cast<IAutoShape*>(shape))->GetTextFrame()->SetText(text.c_str());
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;

}
