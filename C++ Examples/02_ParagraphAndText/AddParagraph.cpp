#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"AddParagraph.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	//Set background image
	std::wstring ImageFile = DataPath"bg.png";
	RectangleF* rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Append a new shape
	IAutoShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(50, 70, 620, 150));
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());

	//Set the alignment of paragraph
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Left);
	//Set the indent of paragraph
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetIndent(50);
	//Set the linespacing of paragraph
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetLineSpacing(150);
	//Set the text of paragraph
	shape->GetTextFrame()->SetText(L"This powerful component suite contains the most up-to-date versions of all .NET WPF Silverlight components offered by E-iceblue.");

	//Set the Font
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Arial Rounded MT Bold"));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
