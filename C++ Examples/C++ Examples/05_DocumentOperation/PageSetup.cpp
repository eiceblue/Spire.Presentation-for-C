#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"PageSetup.pptx";
	std::wstring outputFile = OutputPath"PageSetup.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	//ppt->LoadFromFile(inputFile.c_str());

	//Set the size of slides
	ppt->GetSlideSize()->SetSize(new SizeF(600, 600));
	ppt->GetSlideSize()->SetOrientation(SlideOrienation::Portrait);
	ppt->GetSlideSize()->SetType(SlideSizeType::Custom);

	//Set background image
	std::wstring ImageFile = DataPath"bg.png";
	RectangleF* rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Append new shape
	RectangleF* rec = new RectangleF(ppt->GetSlideSize()->GetSize()->GetWidth() / 2 - 200, 150, 400, 200);
	IAutoShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);

	//Add text to shape
	shape->AppendTextFrame(L"The sample demonstrates how to set slide size.");

	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Myriad Pro"));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetFontHeight(24);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::FromArgb(36, 64, 97));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;

}
