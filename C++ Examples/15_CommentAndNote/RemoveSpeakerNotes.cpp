#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_5.pptx";
	std::wstring outputFile = OutputPath"RemoveSpeakerNotes.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Remove the first speak note.
	slide->GetNotesSlide()->GetNotesTextFrame()->GetParagraphs()->RemoveAt(1);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
