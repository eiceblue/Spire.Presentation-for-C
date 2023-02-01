#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SetAlignmentInTable.pptx";
	std::wstring outputFile = OutputPath"SetAlignmentInTable.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	ITable* table = nullptr;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			table = dynamic_cast<ITable*>(shape);

			//Horizontal Alignment
			//Set the horizontal alignment for the cells in first column 
			table->GetItem(0, 1)->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Left);
			table->GetItem(0, 2)->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Center);
			table->GetItem(0, 3)->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Right);
			table->GetItem(0, 4)->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Justify);

			//Vertical Alignment
			//Set the vertical alignment for the cells in second column 
			table->GetItem(1, 1)->SetTextAnchorType(TextAnchorType::Top);
			table->GetItem(1, 2)->SetTextAnchorType(TextAnchorType::Center);
			table->GetItem(1, 3)->SetTextAnchorType(TextAnchorType::Bottom);
			table->GetItem(1, 4)->SetTextAnchorType(TextAnchorType::None);

			//Both orientaions
			//Set the both horizontal and vertical alignment for the cells in the third column 
			table->GetItem(2, 1)->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Left);
			table->GetItem(2, 1)->SetTextAnchorType(TextAnchorType::Top);

			table->GetItem(2, 2)->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Right);
			table->GetItem(2, 2)->SetTextAnchorType(TextAnchorType::Center);

			table->GetItem(2, 3)->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Justify);
			table->GetItem(2, 3)->SetTextAnchorType(TextAnchorType::Bottom);

			table->GetItem(2, 4)->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Center);
			table->GetItem(2, 4)->SetTextAnchorType(TextAnchorType::Top);
		}
	}

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
