#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AddHyperlinkToText.pptx";
	std::wstring outputFile = OutputPath"AddHyperlinkToText.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Find the text we want to add link to it.
	IAutoShape* shape = dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	TextParagraph* tp = shape->GetTextFrame()->GetTextRange()->GetParagraph();
	std::wstring temp = tp->GetText();

	//Split the original text.
	std::wstring textToLink = L"Spire.Presentation";
	std::wstring::size_type pos = temp.find(textToLink);
	std::wstring strSplit = temp.substr(0, pos);

	//Clear all text.
	tp->GetTextRanges()->Clear();

	//Add new text.
	TextRange* tr = new TextRange(strSplit.c_str());
	tp->GetTextRanges()->Append(tr);

	//Add the hyperlink.
	tr = new TextRange(textToLink.c_str());
	tr->GetClickAction()->SetAddress(L"http://www.e-iceblue.com");
	tp->GetTextRanges()->Append(tr);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
