#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"LinkToASpecificSlide.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	//Append a slide to it.
	ppt->GetSlides()->Append();

	//Add a shape to the second slide.
	IAutoShape* shape = ppt->GetSlides()->GetItem(1)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(10, 50, 200, 50));
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetLine()->SetFillType(FillFormatType::None);
	shape->GetTextFrame()->SetText(L"Jump to the first slide");

	//Create a hyperlink based on the shape and the text on it, linking to the first slide.
	ClickHyperlink* hyperlink = new ClickHyperlink(ppt->GetSlides()->GetItem(0));
	shape->SetClick(hyperlink);
	shape->GetTextFrame()->GetTextRange()->SetClickAction(hyperlink);

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
