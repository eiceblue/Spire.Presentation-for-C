#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ToSVG.pptx";
	std::wstring outputFile = OutputPath"SVG/ToSVG/";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Retain the notes while converting PowerPoint file to svg file.
	ppt->SetIsNoteRetained(true);
	SlideCollection* slides = ppt->GetSlides();
	for (int i = 0; i < slides->GetCount(); i++)
	{
		Stream* svg = slides->GetItem(i)->SaveToSVG();
		svg->Save((outputFile + L"ToSVG-" + to_wstring(i) + L".svg").c_str());
		svg->Dispose();
	}
	delete ppt;

}
