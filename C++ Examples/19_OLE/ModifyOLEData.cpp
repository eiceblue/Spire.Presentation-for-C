#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ModifyOLEData.pptx";
	std::wstring outputFile = OutputPath"ModifyOLEData.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	//Load document from disk
	ppt->LoadFromFile(inputFile.c_str());

	//Loop through the slides and shapes
	SlideCollection* slides = ppt->GetSlides();
	for (int i = 0; i < slides->GetCount(); i++)
	{
		ISlide* slide = slides->GetItem(i);
		for (int j = 0; j < slide->GetShapes()->GetCount(); j++)
		{
			IShape* shape = slide->GetShapes()->GetItem(j);
			if (dynamic_cast<Spire::Presentation::IOleObject*>(shape) != nullptr)
			{
				//Find OLE object
				Spire::Presentation::IOleObject* oleObject = dynamic_cast<Spire::Presentation::IOleObject*>(shape);

				Stream* stream = oleObject->GetDataStream();
				//Get its data and write to file
				if (wcscmp(oleObject->GetProgId(), L"PowerPoint.Show.12") == 0)
				{
					//Load the PPT stream
					Presentation* ppt = new Presentation();
					ppt->LoadFromStream(stream, FileFormat::Auto);
					//Append an image in slide
					std::wstring inputFile1 = DataPath"Logo.png";
					ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, inputFile1.c_str(), new RectangleF(50, 50, 100, 100));
					Stream* stream2 = new Stream();
					ppt->SaveToFile(stream2, FileFormat::Pptx2013);
					stream2->SetPosition(0);
					//Modify the data
					oleObject->SetDataStream(stream2);
				}
			}
		}
	}

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
