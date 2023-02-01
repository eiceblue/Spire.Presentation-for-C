
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ChartSample4.pptx";
	std::wstring outputFile = OutputPath"FillPictureInChartMarker.pptx";

	///Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	IChart* chart = dynamic_cast<IChart*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	Stream* stream = new Stream(DataPath"Logo.png");
	IImageData* IImage = ppt->GetImages()->Append(stream);

	//Create a ChartDataPoint object and specify the index
	ChartDataPoint* dataPoint = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	dataPoint->SetIndex(0);

	//Fill picture in marker
	dataPoint->GetMarkerFill()->GetFill()->SetFillType(FillFormatType::Picture);
	dataPoint->GetMarkerFill()->GetFill()->GetPictureFill()->GetPicture()->SetEmbedImage(IImage);

	//Set marker size
	dataPoint->SetMarkerSize(20);

	//Add the data point in series
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(dataPoint);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
