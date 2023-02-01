#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile_1 = DataPath"TextTemplate.pptx";
	std::wstring inputFile_2 = DataPath"CopyParagraph.pptx";
	std::wstring outputFile = OutputPath"CopyParagraphToAnotherPPT.pptx";

	//Load the source file
	Presentation* ppt1 = new Presentation();
	ppt1->LoadFromFile(inputFile_1.c_str());

	//Get the text from the first shape on the first slide
	IShape* sourceshp = ppt1->GetSlides()->GetItem(0)->GetShapes()->GetItem(0);
	std::wstring text = (dynamic_cast<IAutoShape*>(sourceshp))->GetTextFrame()->GetText();

	//Load the target file
	Presentation* ppt2 = new Presentation();
	ppt2->LoadFromFile(inputFile_2.c_str());

	//Get the first shape on the first slide from the target file
	IShape* destshp = ppt2->GetSlides()->GetItem(0)->GetShapes()->GetItem(0);

	//Add the text to the target file
	std::wstring text2 = (dynamic_cast<IAutoShape*>(destshp))->GetTextFrame()->GetText();
	(dynamic_cast<IAutoShape*>(destshp))->GetTextFrame()->SetText((text2 + L"\n\n" + text).c_str());

	//Save the document
	ppt2->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt2;

}
