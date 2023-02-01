#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile_1 = DataPath"ChangeSlidePosition.pptx";
	std::wstring inputFile_2 = DataPath"PPTSample_N.pptx";
	std::wstring outputFile = OutputPath"ClonePPTAtEndOfAnother.pptx";

	//Load source document from disk
	Presentation* sourcePPT = new Presentation();
	sourcePPT->LoadFromFile(inputFile_1.c_str());

	//Load destination document from disk
	Presentation* destPPT = new Presentation();
	destPPT->LoadFromFile(inputFile_2.c_str());

	//Loop through all slides of source document
	for (int l = 0; l < sourcePPT->GetSlides()->GetCount(); l++)
	{
		ISlide* slide = sourcePPT->GetSlides()->GetItem(l);
		//Append the slide at the end of destination document
		destPPT->GetSlides()->Append(slide);
	}
	//Save the document
	destPPT->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete destPPT;
}
