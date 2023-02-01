#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"RemoveNoteFromSlides.pptx";
	std::wstring outputFile = OutputPath"RemoveNotesAtSpecificSlide.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Get note slide
	NotesSlide* note = slide->GetNotesSlide();
	//Clear note text
	note->GetNotesTextFrame()->SetText(L"");

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
