
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"TrendlineEquation.pptx";
	std::wstring outputFile = OutputPath"ChangesForTrendLineEquation.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get chart on the first slide
	IChart* chart = dynamic_cast<IChart*>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Get the first trendline 
	ITrendlines* trendline = chart->GetSeries()->GetItem(0)->GetTrendLines()[0];

	//Change font size for trendline Equation text
	ParagraphCollection* ps = trendline->GetTrendLineLabel()->GetTextFrameProperties()->GetParagraphs();
	for (int i = 0; i < ps->GetCount(); i++)
	{
		TextParagraph* para = ps->GetItem(i);
		para->GetDefaultCharacterProperties()->SetFontHeight(20);
		for (int j = 0; j < para->GetTextRanges()->GetCount(); j++)
		{
			para->GetTextRanges()->GetItem(j)->SetFontHeight(20);
		}
	}

	//Change position for trendline Equation
	trendline->GetTrendLineLabel()->SetOffsetX(-0.1f);
	trendline->GetTrendLineLabel()->SetOffsetY(-0.05f);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
