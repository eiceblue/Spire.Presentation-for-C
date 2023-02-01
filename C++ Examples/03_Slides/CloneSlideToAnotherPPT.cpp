#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile_1 = DataPath"CloneSlideToAnotherPPT-1.pptx";
	std::wstring inputFile_2 = DataPath"CloneSlideToAnotherPPT-2.pptx";
	std::wstring outputFile = OutputPath"CloneSlideToAnotherPPT.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile_2.c_str());

	//Load the another document and choose the first slide to be cloned
	Presentation* ppt1 = new Presentation();
	ppt1->LoadFromFile(inputFile_1.c_str());
	ISlide* slide1 = ppt1->GetSlides()->GetItem(0);

	//Insert the slide to the specified index in the source presentation
	int index = 1;
	presentation->GetSlides()->Insert(index, slide1);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
