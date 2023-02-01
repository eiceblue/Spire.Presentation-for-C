
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile_1 = DataPath"Template_Ppt_2.pptx";
	std::wstring inputFile_2 = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"CopyChartBetweenPptFiles.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile_1.c_str());

	//Get chart on the first slide
	IChart* chart = dynamic_cast<IChart*>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Create a PPT document
	Presentation* presentation2 = new Presentation();
	//Load the file from disk.
	presentation2->LoadFromFile(inputFile_2.c_str());

	//Copy chart from the first document to the second document.
	presentation2->GetSlides()->Append();
	presentation2->GetSlides()->GetItem(1)->GetShapes()->CreateChart(chart, new RectangleF(100, 100, 500, 300), -1);

	//Save to file.
	presentation2->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
