
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"FormatChartDataLabels.pptx";
	std::wstring outputFile = OutputPath"FormatChartDataLabels.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	IChart* chart = dynamic_cast<IChart*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Get the chart series
	ChartSeriesFormatCollection* sers = chart->GetSeries();

	//Initialize four instances of series label and set parameters of each label
	ChartDataLabel* cd1 = sers->GetItem(0)->GetDataLabels()->Add();
	cd1->SetPercentageVisible(true);
	cd1->GetTextFrame()->SetText(L"Custom Datalabel1");
	cd1->GetTextFrame()->GetTextRange()->SetFontHeight(12);
	cd1->GetTextFrame()->GetTextRange()->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));
	cd1->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
	cd1->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Green);

	ChartDataLabel* cd2 = sers->GetItem(0)->GetDataLabels()->Add();
	cd2->SetPosition(ChartDataLabelPosition::InsideEnd);
	cd2->SetPercentageVisible(true);
	cd2->GetTextFrame()->SetText(L"Custom Datalabel2");
	cd2->GetTextFrame()->GetTextRange()->SetFontHeight(10);
	cd2->GetTextFrame()->GetTextRange()->SetLatinFont(new TextFont(L"Arial"));
	cd2->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
	cd2->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::OrangeRed);

	ChartDataLabel* cd3 = sers->GetItem(0)->GetDataLabels()->Add();
	cd3->SetPosition(ChartDataLabelPosition::Center);
	cd3->SetPercentageVisible(true);
	cd3->GetTextFrame()->SetText(L"Custom Datalabel3");
	cd3->GetTextFrame()->GetTextRange()->SetFontHeight(14);
	cd3->GetTextFrame()->GetTextRange()->SetLatinFont(new TextFont(L"Calibri"));
	cd3->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
	cd3->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Blue);

	ChartDataLabel* cd4 = sers->GetItem(0)->GetDataLabels()->Add();
	cd4->SetPosition(ChartDataLabelPosition::InsideBase);
	cd4->SetPercentageVisible(true);
	cd4->GetTextFrame()->SetText(L"Custom Datalabel4");
	cd4->GetTextFrame()->GetTextRange()->SetFontHeight(12);
	cd4->GetTextFrame()->GetTextRange()->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));
	cd4->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
	cd4->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::OliveDrab);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
