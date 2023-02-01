#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SomePresentation.pptx";
	std::wstring outputFile = OutputPath"ReplaceTextRetentionStyle.pptx";

	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	ppt->GetSlides()->GetItem(0)->ReplaceFirstText(L"use", L"test", true);
	ppt->GetSlides()->GetItem(1)->ReplaceAllText(L"Spire", L"new spire", true);
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
