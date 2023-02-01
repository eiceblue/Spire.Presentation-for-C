#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"PPTSample_N.pptx";
	std::wstring outputFile = OutputPath"SVG/SlideToSVG/SlideToSVG.svg";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Convert the second slide to SVG
	Stream* svg = ppt->GetSlides()->GetItem(0)->SaveToSVG();
	svg->Save(outputFile.c_str());
	svg->Close();

	delete ppt;
}
