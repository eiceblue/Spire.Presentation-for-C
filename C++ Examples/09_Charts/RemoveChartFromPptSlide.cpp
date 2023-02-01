
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_3.pptx";
	std::wstring outputFile = OutputPath"RemoveChartFromPptSlide.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide from the document.
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Remove chart from the slide.
	for (int i = slide->GetShapes()->GetCount() - 1; i >= 0; i--)
	{
		IShape* shape = slide->GetShapes()->GetItem(i);
		if (dynamic_cast<IChart*>(shape) != nullptr)
		{
			slide->GetShapes()->Remove(shape);
		}
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
