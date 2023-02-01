#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Animation.pptx";
	std::wstring outputFile = OutputPath"SetAnimationRepeatType.pptx";

	//Create an instance of presentation document
	Presentation* presentation = new Presentation();
	//Load file
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = presentation->GetSlides()->GetItem(0);
	AnimationEffectCollection* animations = slide->GetTimeline()->GetMainSequence();
	animations->GetItem(0)->GetTiming()->SetAnimationRepeatType(AnimationRepeatType::UtilEndOfSlide);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;

}
