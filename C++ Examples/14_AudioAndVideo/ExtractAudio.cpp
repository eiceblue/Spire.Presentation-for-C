
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"audio.pptx";
	std::wstring outputFile = OutputPath"ExtractAudio.wav";
	//Load a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	ShapeCollection* shapes = presentation->GetSlides()->GetItem(0)->GetShapes();
	int index = 1;
	for (int i = 0; i < shapes->GetCount(); i++)
	{
		if (dynamic_cast<IAudio*>(shapes->GetItem(i)) != nullptr)
		{
			IAudio* audio = dynamic_cast<IAudio*>(shapes->GetItem(i));
			audio->GetData()->SaveToFile(outputFile.c_str());
			index++;
		}
	}
	delete presentation;
}
