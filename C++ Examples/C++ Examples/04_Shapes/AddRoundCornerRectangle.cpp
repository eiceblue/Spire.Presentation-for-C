#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"AddRoundCornerRectagle.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();

	//Set background image
	std::wstring ImageFile = DataPath"bg.png";
	RectangleF* rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(),
		ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Append a round corner rectangle and set its radius
	IAutoShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendRoundRectangle(300, 90, 100, 200, 80);
	//Set the color and fill style of shape
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetLightBlue());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetSkyBlue());
	//Rotate the shape to 90 degree
	shape->SetRotation(90);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
