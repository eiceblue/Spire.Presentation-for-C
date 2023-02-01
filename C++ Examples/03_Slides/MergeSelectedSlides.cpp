#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile_1 = DataPath"InputTemplate.pptx";
	std::wstring inputFile_2 = DataPath"TextTemplate.pptx";
	std::wstring outputFile = OutputPath"MergeSelectedSlides.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();

	//Remove the first slide
	ppt->GetSlides()->RemoveAt(0);

	//Load two PPT files
	Presentation* ppt1 = new Presentation();
	ppt1->LoadFromFile(inputFile_1.c_str());
	Presentation* ppt2 = new Presentation();
	ppt2->LoadFromFile(inputFile_2.c_str());
	//Append all slides in ppt1 to ppt
	for (int i = 0; i < ppt1->GetSlides()->GetCount(); i++)
	{
		ppt->GetSlides()->Append(ppt1->GetSlides()->GetItem(i));
	}

	//Append the second slide in ppt2 to ppt
	ppt->GetSlides()->Append(ppt2->GetSlides()->GetItem(1));

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
