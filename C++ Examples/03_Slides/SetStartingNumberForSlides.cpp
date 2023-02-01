#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ChangeSlidePosition.pptx";
	std::wstring outputFile = OutputPath"SetStartingNumberForSlides.pptx";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT document from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Set 5 as the starting number
	presentation->SetFirstSlideNumber(5);

	//Save file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
