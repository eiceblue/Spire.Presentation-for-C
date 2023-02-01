#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ArrangeShape.pptx";
	std::wstring outputFile = OutputPath"ArrangeShapes.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the specified shape
	IShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0);

	//Bring the shape forward through SetShapeArrange method
	shape->SetShapeArrange(ShapeArrange::BringForward);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
