#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ChangeSlidePosition.pptx";
	std::wstring outputFile = OutputPath"CloneSlideAtTheEnd.pptx";

	//Load PPT document from disk
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	//Append the slide at the end of the document
	presentation->GetSlides()->Append(slide);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
