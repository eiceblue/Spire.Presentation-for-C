
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"SetMasterBackground.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Set the slide background of master
	presentation->GetMasters()->GetItem(0)->GetSlideBackground()->SetType(BackgroundType::Custom);
	presentation->GetMasters()->GetItem(0)->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Solid);
	presentation->GetMasters()->GetItem(0)->GetSlideBackground()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::LightSalmon);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
