#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_5.pptx";
	std::wstring outputFile = OutputPath"RemoveHyperlink.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	ppt->LoadFromFile(inputFile.c_str());

	ISlide* slide = ppt->GetSlides()->GetItem(0);
	//Find the hyperlinks you want to edit.
	IAutoShape* shape = dynamic_cast<IAutoShape*>(slide->GetShapes()->GetItem(0));

	//Set the ClickAction property into null to remove the hyperlink.
	shape->GetTextFrame()->GetTextRange()->SetClickAction(ClickHyperlink::GetNoAction());

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
