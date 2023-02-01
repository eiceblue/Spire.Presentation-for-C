#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Az1.pptx";
	std::wstring outputFile = OutputPath"SetAnchorOfTextFrame.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);
	//Get a shape 
	IAutoShape* shape = dynamic_cast<IAutoShape*>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	shape->GetTextFrame()->SetAnchoringType(TextAnchorType::Bottom);

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;

}
