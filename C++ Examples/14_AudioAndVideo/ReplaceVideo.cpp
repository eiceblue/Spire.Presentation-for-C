
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{

	std::wstring inputFile_1 = DataPath"video.pptx";
	std::wstring inputFile_2 = DataPath"repleaceVido.mp4";
	std::wstring outputFile = OutputPath"ReplaceVideo.pptx";


	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT document from disk.
	presentation->LoadFromFile(inputFile_1.c_str());

	VideoCollection* videos = presentation->GetVideos();

	//Traverse all the slides of PPT file
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		ISlide* slide = presentation->GetSlides()->GetItem(i);
		ShapeCollection* shapes = slide->GetShapes();
		//Traverse all the shapes of slides
		for (int j = 0; j < shapes->GetCount(); j++)
		{
			IShape* shape = shapes->GetItem(j);
			//If shape is IVideo
			if (dynamic_cast<IVideo*>(shape) != nullptr)
			{
				//Replace the video
				IVideo* video = dynamic_cast<IVideo*>(shape);
				//Load the video document from disk.
				Stream* videoStream = new Stream(inputFile_2.c_str());
				VideoData* videoData = videos->Append(videoStream);
				video->SetEmbeddedVideoData(videoData);
			}
		}
	}
	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
