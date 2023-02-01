#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"AddLineToSlide.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	//Add a line in the slide
	IAutoShape* line = slide->GetShapes()->AppendShape(ShapeType::Line, new RectangleF(50, 100, 300, 0));

	//Set color of the line
	line->GetShapeStyle()->GetLineColor()->SetColor(Color::GetRed());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
