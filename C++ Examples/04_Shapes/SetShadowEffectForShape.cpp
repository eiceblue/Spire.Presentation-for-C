#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"SetShadowEffectForShape.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();

	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Set background image
	std::wstring ImageFile = DataPath"bg.png";
	RectangleF* rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	slide->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	slide->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Add a shape to slide.
	RectangleF* rect1 = new RectangleF(200, 150, 300, 120);
	IAutoShape* shape = slide->GetShapes()->AppendShape(ShapeType::Rectangle, rect1);
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetLightBlue());
	shape->GetLine()->SetFillType(FillFormatType::None);
	shape->GetTextFrame()->SetText(L"This demo shows how to apply shadow effect to shape.");
	shape->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());

	//Create an inner shadow effect through InnerShadowEffect object. 
	InnerShadowEffect* innerShadow = new InnerShadowEffect();
	innerShadow->SetBlurRadius(20);
	innerShadow->SetDirection(0);
	innerShadow->SetDistance(0);
	innerShadow->GetColorFormat()->SetColor(Color::GetBlack());

	//Apply the shadow effect to shape.
	shape->GetEffectDag()->SetInnerShadowEffect(innerShadow);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
