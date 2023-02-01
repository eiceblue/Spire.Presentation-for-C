#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"CreateSmartArtShape.pptx";
	std::wstring outputFile = OutputPath"CreateSmartArtShape.pptx";

	// Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	ISmartArt* sa = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendSmartArt(200, 60, 300, 300, SmartArtLayoutType::Gear);

	//Set type and color of smartart
	sa->SetStyle(SmartArtStyleType::SubtleEffect);
	sa->SetColorStyle(SmartArtColorType::GradientLoopAccent3);

	//Remove all shapes
	int i = sa->GetNodes()->GetCount();
	while (i > 0)
	{
		sa->GetNodes()->RemoveNode(0);
		i--;
	}
	//Add two custom shapes with text
	ISmartArtNode* node = sa->GetNodes()->AddNode();
	sa->GetNodes()->GetItem(0)->GetTextFrame()->SetText(L"aa");
	node = sa->GetNodes()->AddNode();
	node->GetTextFrame()->SetText(L"bb");
	node->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
	node->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Black);

	//Save and launch the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
