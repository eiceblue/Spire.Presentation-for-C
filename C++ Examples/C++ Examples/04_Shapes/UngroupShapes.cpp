#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"GroupShapes.pptx";
	std::wstring outputFile = OutputPath"UngroupShapes.pptx";

	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	GroupShape* groupShape = dynamic_cast<GroupShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	//Ungroup the shapes
	ppt->GetSlides()->GetItem(0)->Ungroup(groupShape);
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
