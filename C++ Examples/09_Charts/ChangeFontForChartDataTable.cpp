
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ChartSample2.pptx";
	std::wstring outputFile = OutputPath"ChangeFontSizeForChartDataTable.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get chart on the first slide
	IChart* chart = dynamic_cast<IChart*>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	chart->SetHasDataTable(true);

	//Add a new paragraph in data table
	TextParagraph* tp = new TextParagraph();
	chart->GetChartDataTable()->GetText()->GetParagraphs()->Append(tp);
	//Change the font size
	chart->GetChartDataTable()->GetText()->GetParagraphs()->GetItem(0)
		->GetDefaultCharacterProperties()->SetFontHeight(15);
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
