#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Az2.pptx";
	std::wstring outputFile = OutputPath"SetParagraphFont.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	//Access the first and second placeholder in the slide and typecasting it as AutoShape
	ITextFrameProperties* tf1 = (dynamic_cast<IAutoShape*>(slide->GetShapes()->GetItem(0)))->GetTextFrame();
	ITextFrameProperties* tf2 = (dynamic_cast<IAutoShape*>(slide->GetShapes()->GetItem(1)))->GetTextFrame();

	// Access the first Paragraph
	TextParagraph* para1 = tf1->GetParagraphs()->GetItem(0);
	TextParagraph* para2 = tf2->GetParagraphs()->GetItem(0);

	//Justify the paragraph
	para2->SetAlignment(TextAlignmentType::Justify);

	//Access the first text range
	TextRange* textRange1 = para1->GetFirstTextRange();
	TextRange* textRange2 = para2->GetFirstTextRange();

	//Define new fonts
	TextFont* fd1 = new TextFont(L"Elephant");
	TextFont* fd2 = new TextFont(L"Castellar");

	// Assign new fonts to text range
	textRange1->SetLatinFont(fd1);
	textRange2->SetLatinFont(fd2);

	// Set font to Bold
	textRange1->GetFormat()->SetIsBold(TriState::True);
	textRange2->GetFormat()->SetIsBold(TriState::False);

	// Set font to Italic
	textRange1->GetFormat()->SetIsItalic(TriState::False);
	textRange2->GetFormat()->SetIsItalic(TriState::True);

	// Set font color
	textRange1->GetFill()->SetFillType(FillFormatType::Solid);
	textRange1->GetFill()->GetSolidColor()->SetColor(Color::GetPurple());
	textRange2->GetFill()->SetFillType(FillFormatType::Solid);
	textRange2->GetFill()->GetSolidColor()->SetColor(Color::GetPeru());

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;

}
