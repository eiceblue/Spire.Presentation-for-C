#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_7.pptx";
	std::wstring outputFile = OutputPath"ResetPositionOfDateTimeAndSlideNumber.pptx";

	//Create a PowerPoint document.
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide from the sample document.
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
	{
		IShape* shapeToMove = slide->GetShapes()->GetItem(s);
		//Reset the position of the slide number to the left.
		std::wstring temp = shapeToMove->GetName();
		std::wstring::size_type pos = temp.find(L"Slide Number Placeholder");
		std::wstring temp1 = shapeToMove->GetName();
		std::wstring::size_type pos1 = temp.find(L"Date Placeholder");
		if (pos != string::npos)
		{
			shapeToMove->SetLeft(0);
		}
		else if (pos1 != string::npos)
		{
			//Reset the position of the date time to the center.
			shapeToMove->SetLeft(presentation->GetSlideSize()->GetSize()->GetWidth() / 2);

			//Reset the date time display style.
			(dynamic_cast<IAutoShape*>(shapeToMove))->GetTextFrame()->GetTextRange()->GetParagraph()->SetText((DateTime::GetNow())->ToString());
			(dynamic_cast<IAutoShape*>(shapeToMove))->GetTextFrame()->SetIsCentered(true);
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
