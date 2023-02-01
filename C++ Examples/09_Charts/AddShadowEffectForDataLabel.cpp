
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_3.pptx";
	std::wstring outputFile = OutputPath"AddShadowEffectForDataLabel.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the column chart on the first slide and set chart title.
	IChart* chart = dynamic_cast<IChart*>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Add a data label to the first chart series.
	ChartDataLabelCollection* dataLabels = chart->GetSeries()->GetItem(0)->GetDataLabels();
	ChartDataLabel* Label = dataLabels->Add();
	Label->SetLabelValueVisible(true);

	//Add outer shadow effect to the data label.
	Label->GetEffect()->SetOuterShadowEffect(new OuterShadowEffect());

	//Set shadow color.
	Label->GetEffect()->GetOuterShadowEffect()->GetColorFormat()->SetKnownColor(KnownColors::Yellow);

	//Set blur.
	Label->GetEffect()->GetOuterShadowEffect()->SetBlurRadius(5);

	//Set distance.
	Label->GetEffect()->GetOuterShadowEffect()->SetDistance(10);

	//Set angle.
	Label->GetEffect()->GetOuterShadowEffect()->SetDirection(90.0);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
