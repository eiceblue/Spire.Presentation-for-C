
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_2.pptx";
	std::wstring outputFile = OutputPath"CopyChartWithinOnePPT.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get chart on the first slide
	IChart* chart = dynamic_cast<IChart*>(ppt->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Copy the chart from the first slide to the specified location of the second slide within the same document.
	ISlide* slide1 = ppt->GetSlides()->Append();
	slide1->GetShapes()->CreateChart(chart, new RectangleF(100, 100, 500, 300), 0);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
