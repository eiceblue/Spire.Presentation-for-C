#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ChartSample2.pptx";
	std::wstring outputFile = OutputPath"SetChartBorderstyle.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	IChart* chart = dynamic_cast<IChart*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Set border style
	chart->GetLine()->GetFillFormat()->SetFillType(FillFormatType::Solid);
	chart->GetLine()->GetFillFormat()->GetSolidFillColor()->SetKnownColor(KnownColors::Red);
	chart->SetBorderRoundedCorners(true);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
