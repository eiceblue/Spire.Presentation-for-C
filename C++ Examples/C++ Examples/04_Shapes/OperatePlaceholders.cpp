#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"OperatePlaceholders.pptx";
	std::wstring outputFile = OutputPath"OperatePlaceholders.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Operate placeholders
	for (int j = 0; j < presentation->GetSlides()->GetCount(); j++)
	{
		ISlide* slide = dynamic_cast<ISlide*>(presentation->GetSlides()->GetItem(j));

		for (int i = 0; i < slide->GetShapes()->GetCount(); i++)
		{
			IShape* shape = slide->GetShapes()->GetItem(i);
			switch (shape->GetPlaceholder()->GetType())
			{
			case PlaceholderType::Media:
				shape->InsertVideo(DataPath"Video.mp4");
				break;

			case PlaceholderType::Picture:
				shape->InsertPicture(DataPath"E-iceblueLogo.png");
				break;

			case PlaceholderType::Chart:
				shape->InsertChart(ChartType::ColumnClustered);
				break;

			case PlaceholderType::Table:
				shape->InsertTable(3, 2);
				break;

			case PlaceholderType::Diagram:
				shape->InsertSmartArt(SmartArtLayoutType::BasicBlockList);
				break;
			}
		}
	}
	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
