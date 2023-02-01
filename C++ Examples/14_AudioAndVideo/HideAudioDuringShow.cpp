
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"audio.pptx";
	std::wstring outputFile = OutputPath"HideAudioDuringShow.pptx";

	//Load a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	//Hide Audio during show
	for (int i = 0; i < slide->GetShapes()->GetCount(); i++)
	{
		if (dynamic_cast<IAudio*>(slide->GetShapes()->GetItem(i)) != nullptr)
		{
			IAudio* audio = dynamic_cast<IAudio*>(slide->GetShapes()->GetItem(i));
			audio->SetHideAtShowing(true);
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

	delete presentation;
}
