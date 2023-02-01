#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"AddLineWithTwoPoints.pptx";

	Presentation* ppt = new Presentation();
	//Get the first slide
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Add line with two points
	IAutoShape* line = slide->GetShapes()->AppendShape(ShapeType::Line, new PointF(50, 50), new PointF(150, 150));
	line->GetShapeStyle()->GetLineColor()->SetColor(Color::GetRed());
	line = slide->GetShapes()->AppendShape(ShapeType::Line, new PointF(150, 150), new PointF(250, 50));
	line->GetShapeStyle()->GetLineColor()->SetColor(Color::GetBlue());
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
