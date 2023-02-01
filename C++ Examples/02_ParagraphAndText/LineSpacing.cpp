#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Az.pptx";
	std::wstring outputFile = OutputPath"LineSpacing.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);
	//Add a shape 
	IAutoShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(50, 100, presentation->GetSlideSize()->GetSize()->GetWidth() - 100, 300));
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetTextFrame()->GetParagraphs()->Clear();

	//Add text
	shape->AppendTextFrame(L"Spire.Presentation for .NET is a professional PowerPoint® compatible API that enables developers tocreate, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform.From Spire.Presentation v 3.7.5, Spire.Presentation starts to support .NET Core, .NET standard.");//Set font and color of text
	TextRange* textRange = shape->GetTextFrame()->GetTextRange();
	textRange->GetFill()->SetFillType(FillFormatType::Solid);
	textRange->GetFill()->GetSolidColor()->SetColor(Color::GetBlueViolet());
	textRange->SetFontHeight(20);
	textRange->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));

	//Set properties of paragraph
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetSpaceBefore(100);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetSpaceAfter(100);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetLineSpacing(150);

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;

}
