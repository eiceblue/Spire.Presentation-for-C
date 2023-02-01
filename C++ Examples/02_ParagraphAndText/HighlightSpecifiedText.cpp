#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SomePresentation.pptx";
	std::wstring outputFile = OutputPath"HighlightSpecifiedText.pptx";

	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the specified shape
	IAutoShape* shape = dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));

	TextHighLightingOptions* options = new TextHighLightingOptions();
	options->SetWholeWordsOnly(true);
	options->SetCaseSensitive(true);

	shape->GetTextFrame()->HighLightText(L"Spire", Color::GetYellow(), options);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
