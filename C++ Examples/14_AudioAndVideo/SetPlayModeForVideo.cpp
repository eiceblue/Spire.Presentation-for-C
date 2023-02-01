
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_8.pptx";
	std::wstring outputFile = OutputPath"SetPlayModeForVideo.pptx";
	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT document from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Find the video by looping through all the slides and set its play mode as auto.
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		ISlide* slide = presentation->GetSlides()->GetItem(i);
		ShapeCollection* shapes = slide->GetShapes();
		for (int j = 0; j < shapes->GetCount(); j++)
		{
			IShape* shape = shapes->GetItem(j);
			//If shape is IVideo
			if (dynamic_cast<IVideo*>(shape) != nullptr)
			{
				//Replace the video
				IVideo* video = dynamic_cast<IVideo*>(shape);
				video->SetPlayMode(VideoPlayMode::Auto);
			}
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
