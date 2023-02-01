#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SmartArtNode.pptx";
	std::wstring outputFile = OutputPath"ChangeHyperlinkColor.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Get the theme of the slide
	Theme* theme = slide->GetTheme();

	//Change the color of hyperlink to red
	theme->GetColorScheme()->GetHyperlinkColor()->SetColor(Color::GetRed());

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
