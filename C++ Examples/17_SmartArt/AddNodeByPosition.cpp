#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AddSmartArtNode2.pptx";
	std::wstring outputFile = OutputPath"AddNodeByPosition.pptx";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ISmartArt*>(shape) != nullptr)
		{
			//Get the SmartArt and collect nodes
			ISmartArt* smartArt = dynamic_cast<ISmartArt*>(shape);

			int position = 0;
			//Add a new node at specific position
			ISmartArtNode* node = smartArt->GetNodes()->AddNodeByPosition(position);
			//Add text and set the text style 
			node->GetTextFrame()->SetText(L"New Node");
			node->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
			node->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Red);

			//Get a node
			node = smartArt->GetNodes()->GetItem(1);
			position = 1;
			//Add a new child node at specific position
			ISmartArtNode* childNode = node->GetChildNodes()->AddNodeByPosition(position);
			//Add text and set the text style 
			node->GetTextFrame()->SetText(L"New child node");
			node->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
			node->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Blue);
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
