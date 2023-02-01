#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ExtractImage.pptx";
	std::wstring outputFile = OutputPath"ChangeImageSize.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	SlideCollection* slides = ppt->GetSlides();
	float scale = 0.5f;
	for (int i = 0; i < slides->GetCount(); i++)
	{
		ShapeCollection* shapes = slides->GetItem(i)->GetShapes();
		for (int j = 0; j < shapes->GetCount(); j++)
		{
			IShape* shape = shapes->GetItem(j);
			if (dynamic_cast<IEmbedImage*>(shape) != nullptr)
			{
				IEmbedImage* image = dynamic_cast<IEmbedImage*>(shape);
				image->SetWidth(image->GetWidth() * scale);
				image->SetHeight(image->GetHeight() * scale);
			}
		}
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
