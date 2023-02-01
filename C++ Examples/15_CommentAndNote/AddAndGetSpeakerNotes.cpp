#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile_px = OutputPath"AddAndGetSpeakerNotes.pptx";
	std::wstring outputFile_txt = OutputPath"AddAndGetSpeakerNotes.txt";

	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide and in the PowerPoint document.
	ISlide* slide = presentation->GetSlides()->GetItem(0);

	//Get the NotesSlide in the first slide,if there is no notes, we need to add it firstly.
	NotesSlide* ns = slide->GetNotesSlide();
	if (ns == nullptr)
	{
		ns = slide->AddNotesSlide();
	}
	//Add the text string as the notes.
	ns->GetNotesTextFrame()->SetText(L"Speak notes added by Spire.Presentation");
	wofstream desFile(outputFile_txt, ios::out);
	//Get the speaker notes and save to txt file.
	desFile << "The speaker notes added by Spire.Presentation is: " << ns->GetNotesTextFrame()->GetText() << endl;
	desFile.close();

	//Save to PowerPoint file.
	presentation->SaveToFile(outputFile_px.c_str(), FileFormat::Pptx2013);

	delete presentation;
}
