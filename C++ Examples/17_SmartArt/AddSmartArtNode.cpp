#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AddSmartArtNode.pptx";
	std::wstring outputFile = OutputPath"AddSmartArtNode.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the SmartArt
	ISmartArt* sa = dynamic_cast<ISmartArt*>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Add a node
	ISmartArtNode* node = sa->GetNodes()->AddNode();
	//Add text and set the text style 
	node->GetTextFrame()->SetText(L"AddText");
	node->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
	node->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::HotPink);

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
