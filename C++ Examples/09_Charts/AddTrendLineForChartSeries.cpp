
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_2.pptx";
	std::wstring outputFile = OutputPath"AddTrendLineForChartSeries.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the target chart, add trendline for the first data series of the chart and specify the trendline type.
	IChart* chart = dynamic_cast<IChart*>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	ITrendlines* it = chart->GetSeries()->GetItem(0)->AddTrendLine(TrendlinesType::Linear);

	//Set the trendline properties to determine what should be displayed.
	it->SetDisplayEquation(false);
	it->SetDisplayRSquaredValue(false);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
