#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"TwoColumns.pptx";
	std::wstring outputFile = OutputPath"SetRightToLeftColumns.pptx";

	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	//Get the second shape
	IAutoShape* shape = dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));
	//Set columns style to right-to-left
	shape->GetTextFrame()->SetRightToLeftColumns(true);
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
