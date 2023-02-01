#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ChangeSlideLayout.pptx";
	std::wstring outputFile = OutputPath"ChangeSlideLayout.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Change the layout of slide
	presentation->GetSlides()->GetItem(1)->SetLayout(presentation->GetMasters()->GetItem(0)->GetLayouts()->GetItem(4));

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
