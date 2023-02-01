#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"HideSlide.pptx";
	std::wstring outputFile = OutputPath"HideSlide.pptx";

	//Create a PPT document and load PPT file from disk
	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Hide the second slide
	ppt->GetSlides()->GetItem(1)->SetHidden(true);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
