#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"OverlappingShapes.pptx";
	std::wstring outputFile = OutputPath"ReorderOverlappingShapes.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first shape of the first slide
	IShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0);
	//Change the shape's zorder
	ppt->GetSlides()->GetItem(0)->GetShapes()->ZOrder(1, shape);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
