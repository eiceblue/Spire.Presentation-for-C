
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"GroupTwoLevelAxisLabels.pptx";
	std::wstring outputFile = OutputPath"GroupTwoLevelAxisLabels.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	IChart* chart = dynamic_cast<IChart*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Get the category axis from the chart.
	IChartAxis* chartAxis = chart->GetPrimaryCategoryAxis();

	//Group the axis labels that have the same first-level label.
	if (chartAxis->GetHasMultiLvlLbl())
	{
		chartAxis->SetIsMergeSameLabel(true);
	}

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
