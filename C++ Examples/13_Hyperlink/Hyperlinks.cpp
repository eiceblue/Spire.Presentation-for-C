#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"Hyperlinks.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	//Set background Image
	std::wstring ImageFile = DataPath"bg.png";
	RectangleF* rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);

	//Add new shape to PPT document
	RectangleF* rec = new RectangleF(ppt->GetSlideSize()->GetSize()->GetWidth() / 2 - 255, 120, 500, 280);
	IAutoShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetLine()->SetWidth(0);

	//Add some paragraphs with hyperlinks
	TextParagraph* para1 = new TextParagraph();
	TextRange* tr = new TextRange(L"E-iceblue");
	tr->GetFill()->SetFillType(FillFormatType::Solid);
	tr->GetFill()->GetSolidColor()->SetColor(Color::GetBlue());
	para1->GetTextRanges()->Append(tr);
	para1->SetAlignment(TextAlignmentType::Center);
	shape->GetTextFrame()->GetParagraphs()->Append(para1);
	TextParagraph* tp = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(tp);

	//Add some paragraphs with hyperlinks
	TextParagraph* para2 = new TextParagraph();
	TextRange* tr1 = new TextRange(L"Click to know more about Spire.Presentation.");
	tr1->GetClickAction()->SetAddress(L"http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html");
	para2->GetTextRanges()->Append(tr1);
	shape->GetTextFrame()->GetParagraphs()->Append(para2);
	TextParagraph* tp2 = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(tp2);

	TextParagraph* para3 = new TextParagraph();
	TextRange* tr2 = new TextRange(L"Click to visit E-iceblue Home page.");
	tr2->GetClickAction()->SetAddress(L"https://www.e-iceblue.com/");
	para3->GetTextRanges()->Append(tr2);
	shape->GetTextFrame()->GetParagraphs()->Append(para3);
	TextParagraph* temp3 = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(temp3);

	TextParagraph* para4 = new TextParagraph();
	TextRange* tr3 = new TextRange(L"Click to go to the forum to raise questions.");
	tr3->GetClickAction()->SetAddress(L"https://www.e-iceblue.com/forum/components-f5.html");
	para4->GetTextRanges()->Append(tr3);
	shape->GetTextFrame()->GetParagraphs()->Append(para4);
	TextParagraph* temp4 = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(temp4);

	TextParagraph* para5 = new TextParagraph();
	TextRange* tr4 = new TextRange(L"Click to contact our sales team via email.");
	tr4->GetClickAction()->SetAddress(L"mailto:sales@e-iceblue.com");
	para5->GetTextRanges()->Append(tr4);
	shape->GetTextFrame()->GetParagraphs()->Append(para5);
	TextParagraph* temp5 = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(temp5);

	TextParagraph* para6 = new TextParagraph();
	TextRange* tr5 = new TextRange(L"Click to contact our support team via email.");
	tr5->GetClickAction()->SetAddress(L"mailto:support@e-iceblue.com");
	para6->GetTextRanges()->Append(tr5);
	shape->GetTextFrame()->GetParagraphs()->Append(para6);

	for (int i = 0; i < shape->GetTextFrame()->GetParagraphs()->GetCount(); i++)
	{
		TextParagraph* temp6 = shape->GetTextFrame()->GetParagraphs()->GetItem(i);
		if (temp6->GetTextRanges()->GetCount() > 0)
		{
			temp6->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));
			temp6->GetTextRanges()->GetItem(0)->SetFontHeight(20);
		}
	}

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;

}
