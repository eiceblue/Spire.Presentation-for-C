#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"PPTHasHeader.pptx";
	std::wstring outputFile = OutputPath"ManageNoteMasterHeaderFooter.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Set the note Masters header and footer
	INoteMasterSlide* noteMasterSlide = ppt->GetNotesMaster();
	if (noteMasterSlide != nullptr)
	{
		ShapeCollection* shapes = noteMasterSlide->GetShapes();
		for (int i = 0; i < shapes->GetCount(); i++)
		{
			IShape* shape = shapes->GetItem(i);
			if (shape->GetPlaceholder() != nullptr)
			{
				if (shape->GetPlaceholder()->GetType() == PlaceholderType::Header)
				{
					(dynamic_cast<IAutoShape*>(shape))->GetTextFrame()->SetText(L"change the header by Spire");
				}
				if (shape->GetPlaceholder()->GetType() == PlaceholderType::Footer)
				{
					(dynamic_cast<IAutoShape*>(shape))->GetTextFrame()->SetText(L"change the footer by Spire");
				}
			}
		}
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
