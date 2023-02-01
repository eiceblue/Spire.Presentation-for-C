#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ChangeSlidePosition.pptx";
	std::wstring outputFile = OutputPath"ChangeSlidePosition.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Move the first slide to the second slide position
	ISlide* slide = presentation->GetSlides()->GetItem(0);
	slide->SetSlideNumber(2);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
