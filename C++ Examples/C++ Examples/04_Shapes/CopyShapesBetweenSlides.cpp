#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"CopyShapesBetweenSlides.pptx";
	std::wstring outputFile = OutputPath"CopyShapesBetweenSlides.pptx";

	//Load the sample document
	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Define the source slide and target slide
	ISlide* sourceSlide = ppt->GetSlides()->GetItem(0);
	ISlide* targetSlide = ppt->GetSlides()->GetItem(1);

	//Copy the first shape from the source slide to the target slide
	targetSlide->GetShapes()->AddShape(sourceSlide->GetShapes()->GetItem(0));

	//Save the document to file 
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
