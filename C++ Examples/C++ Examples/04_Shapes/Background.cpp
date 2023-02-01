#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"Background.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Set background Image
	std::wstring ImageFile = DataPath"backgroundImg.png";
	RectangleF* rect = new RectangleF(0, 0, presentation->GetSlideSize()->GetSize()->GetWidth(), presentation->GetSlideSize()->GetSize()->GetHeight());
	presentation->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);

	//Add title
	RectangleF* rec_title = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 200, 70, 380, 50);
	IAutoShape* shape_title = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec_title);
	shape_title->GetLine()->SetFillType(FillFormatType::None);
	shape_title->GetFill()->SetFillType(FillFormatType::None);
	TextParagraph* para_title = new TextParagraph();
	para_title->SetText(L"Background Sample");
	para_title->SetAlignment(TextAlignmentType::Center);
	para_title->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));
	para_title->GetTextRanges()->GetItem(0)->SetFontHeight(36);
	para_title->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	para_title->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetDarkSlateBlue());
	shape_title->GetTextFrame()->GetParagraphs()->Append(para_title);

	//Add new shape to PPT document
	RectangleF* rec = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 300, 155, 600, 200);
	IAutoShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);
	shape->GetLine()->SetFillType(FillFormatType::None);
	shape->GetFill()->SetFillType(FillFormatType::None);

	TextParagraph* para = new TextParagraph();
	para->SetText(L"Spire.Presentation for .NET support PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio and video on slides. It also support exporting presentation slides to EMF, JPG, TIFF, PDF format etc.");

	para->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	para->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());
	para->GetTextRanges()->GetItem(0)->SetFontHeight(26);
	shape->GetTextFrame()->GetParagraphs()->Append(para);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
