#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Animation.pptx";
	std::wstring outputFile = OutputPath"SetAnimationForAnimateText.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Set the AnimateType as Letter
	ppt->GetSlides()->GetItem(0)->GetTimeline()->GetMainSequence()->GetItem(0)->SetIterateType(AnimateType::Letter);

	//Set the IterateTimeValue for the animate text
	ppt->GetSlides()->GetItem(0)->GetTimeline()->GetMainSequence()->GetItem(0)->SetIterateTimeValue(10);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
