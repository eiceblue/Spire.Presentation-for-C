
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ChartSample2.pptx";
	std::wstring outputFile = OutputPath"ModifyChartCategoryAxis.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	IChart* chart = dynamic_cast<IChart*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Modify the major unit
	chart->GetPrimaryCategoryAxis()->SetIsAutoMajor(false);
	chart->GetPrimaryCategoryAxis()->SetMajorUnit(1);
	chart->GetPrimaryCategoryAxis()->SetMajorUnitScale(ChartBaseUnitType::Months);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
