#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ChangeSlidePosition.pptx";
	std::wstring outputFile = OutputPath"SpecificSlideToPDF.pdf";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the second slide
	ISlide* slide = ppt->GetSlides()->GetItem(1);

	//Save the second slide to PDF
	slide->SaveToFile(outputFile.c_str(), FileFormat::PDF);

	delete ppt;
}
