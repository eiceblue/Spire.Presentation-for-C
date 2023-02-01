#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Conversion.pptx";
	std::wstring outputFile = OutputPath"Image/ToSpecificSizeImage.png";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save the first slide to Image and set the image size to 600*400
	Stream* stream = ppt->GetSlides()->GetItem(0)->SaveAsImage(600, 400);
	stream->Save(outputFile.c_str());

	delete ppt;

}
