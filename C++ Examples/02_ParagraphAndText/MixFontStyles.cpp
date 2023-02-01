#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"FontStyle.pptx";
	std::wstring outputFile = OutputPath"MixFontStyles.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the second shape of the first slide
	IAutoShape* shape = dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));
	//Get the text from the shape
	std::wstring originalText = shape->GetTextFrame()->GetText();

	std::vector<std::wstring> splitArray = { L"Here is testing text. Only a few words are in ",
		L", some are in ", L" color, some are ",L", and some are in ",L"." };

	//Remove the paragraph from TextRange
	TextParagraph* tp = shape->GetTextFrame()->GetTextRange()->GetParagraph();
	tp->GetTextRanges()->Clear();

	//Append normal text that is in front of 'bold' to the paragraph
	TextRange* tr = new TextRange(splitArray[0].c_str());
	tp->GetTextRanges()->Append(tr);
	//Set font style of the text 'bold' as bold
	tr = new TextRange(L"bold");
	tr->SetIsBold(TriState::True);
	tp->GetTextRanges()->Append(tr);

	//Append normal text that is in front of 'red' to the paragraph
	tr = new TextRange(splitArray[1].c_str());
	tp->GetTextRanges()->Append(tr);
	//Set the color of the text 'red' as red
	tr = new TextRange(L"red");
	tr->GetFill()->SetFillType(FillFormatType::Solid);
	tr->GetFormat()->GetFill()->GetSolidColor()->SetColor(Color::GetRed());
	tp->GetTextRanges()->Append(tr);

	//Append normal text that is in front of 'underlined' to the paragraph
	tr = new TextRange(splitArray[2].c_str());
	tp->GetTextRanges()->Append(tr);
	//Underline the text 'undelined'
	tr = new TextRange(L"underlined");
	tr->SetTextUnderlineType(TextUnderlineType::Single);
	tp->GetTextRanges()->Append(tr);

	//Append normal text that is in front of 'bigger font size' to the paragraph
	tr = new TextRange(splitArray[3].c_str());
	tp->GetTextRanges()->Append(tr);
	//Set a large font for the text 'bigger font size'
	tr = new TextRange(L"bigger font size");
	tr->SetFontHeight(35);
	tp->GetTextRanges()->Append(tr);

	//Append other normal text
	tr = new TextRange(splitArray[4].c_str());
	tp->GetTextRanges()->Append(tr);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
