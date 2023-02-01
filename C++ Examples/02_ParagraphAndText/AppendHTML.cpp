#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AppendHTML.pptx";
	std::wstring outputFile = OutputPath"AppendHTML.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	//Add a shape 
	IAutoShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(150, 100, 200, 200));

	//Clear default paragraphs 
	shape->GetTextFrame()->GetParagraphs()->Clear();

	std::wstring code = L"<html><body><p>This is a paragraph</p></body></html>";

	//Append HTML, and generate a paragraph with default style in PPT document.
	shape->GetTextFrame()->GetParagraphs()->AddFromHtml(code.c_str());
	std::wstring codeColor = L"<html><body><p style=\" color:black \">This is a paragraph</p></body></html>";
	//Append HTML with black setting
	shape->GetTextFrame()->GetParagraphs()->AddFromHtml(codeColor.c_str());

	//Add another shape
	IAutoShape* shape1 = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(350, 100, 200, 200));

	//Clear default paragraph 
	shape1->GetTextFrame()->GetParagraphs()->Clear();

	//Change the fill format of shape
	shape1->GetFill()->SetFillType(FillFormatType::Solid);
	shape1->GetFill()->GetSolidColor()->SetColor(Color::GetWhite());

	//Append HTML
	shape1->GetTextFrame()->GetParagraphs()->AddFromHtml(code.c_str());
	TextParagraph* par = shape1->GetTextFrame()->GetParagraphs()->GetItem(0);

	for (int r = 0; r < par->GetTextRanges()->GetCount(); r++) {
		par->GetTextRanges()->GetItem(r)->GetFill()->SetFillType(FillFormatType::Solid);
		par->GetTextRanges()->GetItem(r)->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());
	}

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	//delete ppt;
	delete ppt;

}
