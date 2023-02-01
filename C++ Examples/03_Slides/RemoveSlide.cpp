#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"RemoveSlide.pptx";
	std::wstring outputFile = OutputPath"RemoveSlide.pptx";

	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Remove slide by index
	presentation->GetSlides()->RemoveAt(0);

	//Remove slide by its reference
	ISlide* slide = presentation->GetSlides()->GetItem(1);
	presentation->GetSlides()->Remove(slide);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
