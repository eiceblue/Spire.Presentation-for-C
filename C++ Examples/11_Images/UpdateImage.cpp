#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"UpdateImage.pptx";
	std::wstring outputFile = OutputPath"UpdateImage.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	std::wstring inputFile1 = DataPath"PresentationIcon.png";
	//Append a new image to replace an existing image
	IImageData* image = ppt->GetImages()->Append(new Stream(inputFile1.c_str()));

	//Replace the image which title is "image1" with the new image
	for (int i = 0; i < slide->GetShapes()->GetCount(); i++)
	{
		IShape* shape = slide->GetShapes()->GetItem(i);
		if (dynamic_cast<IEmbedImage*>(shape) != nullptr)
		{
			if (wcscmp(shape->GetAlternativeTitle(), L"image1") == 0)
			{
				(dynamic_cast<IEmbedImage*>(shape))->GetPictureFill()->GetPicture()->SetEmbedImage(image);
			}
		}
	}

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
