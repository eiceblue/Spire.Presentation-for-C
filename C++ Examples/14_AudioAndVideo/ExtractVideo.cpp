
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"video.pptx";
	std::wstring outputFile = OutputPath"ExtractVideo/";
	//Load a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());
	SlideCollection* slides = presentation->GetSlides();

	int index = 0;
	for (int i = 0; i < slides->GetCount(); i++)
	{
		ShapeCollection* shapes = slides->GetItem(i)->GetShapes();
		for (int j = 0; j < shapes->GetCount(); j++)
		{
			if (dynamic_cast<IVideo*>(shapes->GetItem(j)) != nullptr)
			{
				IVideo* audio = dynamic_cast<IVideo*>(shapes->GetItem(j));
				audio->GetEmbeddedVideoData()->SaveToFile((outputFile + L"ExtractVideo_" + to_wstring(index) + L".avi").c_str());
				index++;
			}
		}
	}
	delete presentation;
}
