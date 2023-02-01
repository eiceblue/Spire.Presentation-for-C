#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"SetTextTransparency.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();

	//Set background image
	std::wstring ImageFile = DataPath"bg.png";
	RectangleF* rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Add a shape
	IAutoShape* textboxShape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(100, 100, 300, 120));
	textboxShape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetTransparent());
	textboxShape->GetFill()->SetFillType(FillFormatType::None);

	//Remove default blank paragraphs
	textboxShape->GetTextFrame()->GetParagraphs()->Clear();

	//Add three paragraphs, apply color with different alpha values to text
	int alpha = 55;
	for (int i = 0; i < 3; i++)
	{
		TextParagraph* tp = new TextParagraph();
		textboxShape->GetTextFrame()->GetParagraphs()->Append(tp);
		TextRange* tr = new TextRange(L"Text Transparency");
		textboxShape->GetTextFrame()->GetParagraphs()->GetItem(i)->GetTextRanges()->Append(tr);
		textboxShape->GetTextFrame()->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
		textboxShape->GetTextFrame()->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::FromArgb(alpha, Color::GetPurple()));
		alpha += 100;
	}

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
