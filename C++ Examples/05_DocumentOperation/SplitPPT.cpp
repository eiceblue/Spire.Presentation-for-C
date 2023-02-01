#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"InputTemplate.pptx";
	std::wstring outputFile = OutputPath"SplitPPT/";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Specify the presentation show type as kiosk
	ppt->SetShowType(SlideShowType::Kiosk);

	SlideCollection* slides = ppt->GetSlides();
	for (int i = 0; i < slides->GetCount(); i++)
	{
		std::wstring res = outputFile + L"SplitPPT-" + to_wstring(i) + L".pptx";
		slides->GetItem(i)->SaveToFile(res.c_str(), FileFormat::Pptx2010);
	}
	delete ppt;

}
