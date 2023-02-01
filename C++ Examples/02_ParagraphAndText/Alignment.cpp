#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Alignment.pptx";
	std::wstring outputFile = OutputPath"Alignment.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	ppt->LoadFromFile(inputFile.c_str());

	//Get the related shape and set the text alignment
	IAutoShape* shape = dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Left);
	shape->GetTextFrame()->GetParagraphs()->GetItem(1)->SetAlignment(TextAlignmentType::Center);
	shape->GetTextFrame()->GetParagraphs()->GetItem(2)->SetAlignment(TextAlignmentType::Right);
	shape->GetTextFrame()->GetParagraphs()->GetItem(3)->SetAlignment(TextAlignmentType::Justify);
	shape->GetTextFrame()->GetParagraphs()->GetItem(4)->SetAlignment(TextAlignmentType::None);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;

}
