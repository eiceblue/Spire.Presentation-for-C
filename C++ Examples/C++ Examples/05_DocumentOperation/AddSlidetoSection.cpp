#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Section.pptx";
	std::wstring outputFile = OutputPath"AddSlidetoSection.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Add a new shape to the PPT document
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(200, 50, 300, 100));

	//Create a new section and copy the first slide to it
	Section* NewSection = ppt->GetSectionList()->Append(L"New Section");
	NewSection->Insert(0, ppt->GetSlides()->GetItem(0));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
