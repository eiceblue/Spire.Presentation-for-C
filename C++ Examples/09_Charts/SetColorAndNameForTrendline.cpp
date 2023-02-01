#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SetColorAndNameForTrendline.pptx";
	std::wstring outputFile = OutputPath"SetColorAndNameForTrendline.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	IChart* chart = dynamic_cast<IChart*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Find the first trendline in the chart
	ITrendlines* trendline = dynamic_cast<ITrendlines*>(chart->GetSeries()->GetItem(0)->GetTrendLines()[0]);

	//Set name for trendline
	trendline->SetName(L"trendlineName");

	//Set color for trendline
	trendline->GetLine()->SetFillType(FillFormatType::Solid);
	trendline->GetLine()->GetSolidFillColor()->SetColor(Color::GetRed());

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
