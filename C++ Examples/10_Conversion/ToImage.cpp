#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ToImage.pptx";
	std::wstring outputFile = OutputPath"Image/ToImage/";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save PPT document to images
	SlideCollection* slides = ppt->GetSlides();
	for (int i = 0; i < slides->GetCount(); i++)
	{
		ISlide* slide = slides->GetItem(i);
		Stream* image = slide->SaveAsImage();
		image->Save((outputFile + L"ToImage_img_" + to_wstring(i) + L".png").c_str());
	}

	delete ppt;
}
