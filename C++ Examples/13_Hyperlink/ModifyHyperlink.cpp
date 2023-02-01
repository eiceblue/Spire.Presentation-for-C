#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_5.pptx";
	std::wstring outputFile = OutputPath"ModifyHyperlink.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	ppt->LoadFromFile(inputFile.c_str());

	ISlide* slide = ppt->GetSlides()->GetItem(0);
	//Find the hyperlinks you want to edit.
	IAutoShape* shape = dynamic_cast<IAutoShape*>(slide->GetShapes()->GetItem(0));

	//Edit the link text and the target URL.
	shape->GetTextFrame()->GetTextRange()->GetClickAction()->SetAddress(L"http://www.e-iceblue.com");
	shape->GetTextFrame()->GetTextRange()->SetText(L"E-iceblue");

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
