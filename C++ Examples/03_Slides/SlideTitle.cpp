#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"InputTemplate.pptx";
	std::wstring outputFile = OutputPath"SlideTitle.pptx";

	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	//Get the first slide
	ISlide* slide = ppt->GetSlides()->GetItem(0);
	//Get the title of the first slide
	std::wstring slideTitle = slide->GetTitle();
	//Set the title of the second slide
	ppt->GetSlides()->GetItem(1)->SetTitle(L"Second Slide");
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
