#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AppendSlideWithMasterLayout.pptx";
	std::wstring outputFile = OutputPath"AppendSlideWithMasterLayout.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the master
	IMasterSlide* master = presentation->GetMasters()->GetItem(0);

	//Get master layout slides
	IMasterLayouts* masterLayouts = master->GetLayouts();
	ILayout* layoutSlide = masterLayouts->GetItem(1);

	////Append a rectangle to the layout slide
	IAutoShape* shape = layoutSlide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(10, 50, 100, 80));

	////Add a text into the shape and set the style
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->AppendTextFrame(L"Layout slide 1");
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Arial Black"));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());

	//Append new slide with master layout
	presentation->GetSlides()->Append(presentation->GetSlides()->GetItem(0), master->GetLayouts()->GetItem(1));

	//Another way to append new slide with master layout
	presentation->GetSlides()->Insert(2, presentation->GetSlides()->GetItem(1), master->GetLayouts()->GetItem(1));

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
