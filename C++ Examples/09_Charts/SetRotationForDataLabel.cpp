#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SetRotationForDataLabel.pptx";
	std::wstring outputFile = OutputPath"SetRotationForDataLabel.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	IChart* chart = dynamic_cast<IChart*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Set the rotation angle for the datalabels of first serie
	for (int i = 0; i < chart->GetSeries()->GetItem(0)->GetValues()->GetCount(); i++)
	{
		ChartDataLabel* datalabel = chart->GetSeries()->GetItem(0)->GetDataLabels()->Add();
		datalabel->SetID(i);
		datalabel->SetRotationAngle(45);
	}

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
