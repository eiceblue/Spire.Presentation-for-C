#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"FillShapeWithPattern.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	//Add a rectangle
	RectangleF* rect = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 50, 100, 100, 100);
	IAutoShape* shape = slide->GetShapes()->AppendShape(ShapeType::Rectangle, rect);

	//Set the pattern fill format 
	shape->GetFill()->SetFillType(FillFormatType::Pattern);
	shape->GetFill()->GetPattern()->SetPatternType(PatternFillType::Trellis);
	shape->GetFill()->GetPattern()->GetBackgroundColor()->SetColor(Color::GetDarkGray());
	shape->GetFill()->GetPattern()->GetForegroundColor()->SetColor(Color::GetYellow());

	//Set the fill format of line
	shape->GetLine()->SetFillType(FillFormatType::Solid);
	shape->GetLine()->GetSolidFillColor()->SetColor(Color::GetTransparent());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
