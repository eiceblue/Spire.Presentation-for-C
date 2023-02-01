#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"TextTemplate.pptx";
	std::wstring outputFile = OutputPath"ReplaceText.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	ISlide* slide = ppt->GetSlides()->GetItem(0);

	slide->ReplaceAllText(L"Spire.Presentation for .NET", L"Spire.PPT", false);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
