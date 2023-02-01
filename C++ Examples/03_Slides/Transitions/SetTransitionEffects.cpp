#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SetTransitions.pptx";
	std::wstring outputFile = OutputPath"SetTransitionEffects.pptx";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	// Set effects
	presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->SetType(TransitionType::Cut);
	(dynamic_cast<OptionalBlackTransition*>(presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->GetValue()))->SetFromBlack(true);

	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
