
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_2.pptx";
	std::wstring outputFile = OutputPath"HideOrShowASeriesOfChart.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	IChart* chart = dynamic_cast<IChart*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Hide the first series of the chart.
	chart->GetSeries()->GetItem(0)->SetIsHidden(true);

	//Show the first series of the chart.
	//chart.Series[0].IsHidden = false;

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
