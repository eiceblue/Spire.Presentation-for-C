#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Images.pptx";
	std::wstring outputFile = OutputPath"Image/ExtractImageFromSpecificSlide/";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the pictures on the second slide and save them to image file
	int i = 0;
	ShapeCollection* shapes = ppt->GetSlides()->GetItem(1)->GetShapes();
	//Traverse all shapes in the second slide
	for (int j = 0; j < shapes->GetCount(); j++)
	{
		IShape* s = shapes->GetItem(j);
		//It is the SlidePicture object
		if (dynamic_cast<IEmbedImage*>(s) != nullptr)
		{
			//Save to image
			IEmbedImage* ps = dynamic_cast<IEmbedImage*>(s);
			ps->GetPictureFill()->GetPicture()->GetEmbedImage()->GetImage()->Save((outputFile + L"SlidePic_" + to_wstring(i) + L".png").c_str());
			i++;
		}
		//It is the PictureShape object
		if (dynamic_cast<PictureShape*>(s) != nullptr)
		{
			//Save to image
			PictureShape* ps = dynamic_cast<PictureShape*>(s);
			ps->GetEmbedImage()->GetImage()->Save((outputFile + L"ShapePic_" + to_wstring(i) + L".png").c_str());
			i++;
		}
	}
	delete ppt;

}
