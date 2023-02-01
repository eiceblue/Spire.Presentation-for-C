#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"SaveToStream.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	//Set background image
	std::wstring ImageFile = DataPath"bg.png";
	RectangleF* rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Append new shape
	RectangleF* rec = new RectangleF(50, 100, 600, 150);
	IAutoShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);

	//Add text to shape
	shape->AppendTextFrame(L"This demo shows how to Create PowerPoint file and save it to Stream.");
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetFontHeight(30);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());

	//Save to Stream.
	Stream* stream = new Stream();
	ppt->SaveToFile(stream, FileFormat::Pptx2013);
	stream->Save(outputFile.c_str());
	stream->Close();
	delete ppt;

}
