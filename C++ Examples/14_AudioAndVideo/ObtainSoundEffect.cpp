
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Animation.pptx";
	std::wstring outputFile = OutputPath"ObtainSoundEffect.txt";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Get the audio in a time node
	TimeNodeAudio* audio = slide->GetTimeline()->GetMainSequence()->GetItem(0)->GetTimeNodeAudios().front();

	//Get the properties of the audio, such as sound name, volume or detect if it's mute
	//Save the properties of the audio to Text file
	wofstream desFile(outputFile, ios::out);
	desFile << "SoundName: " << audio->GetSoundName() << endl;
	desFile << "Volume: " << audio->GetVolume() << endl;
	desFile << "IsMute: " << audio->GetIsMute() << endl;

	desFile.close();
}
