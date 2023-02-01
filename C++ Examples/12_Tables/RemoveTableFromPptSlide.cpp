#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"RemoveTableFromPptSlide.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the tables within the PPT document.
	std::vector<IShape*> shape_tems;

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ITable*>(shape) != nullptr)
		{
			//Add new table to table list.
			shape_tems.push_back(shape);
		}
	}
	//Remove all the tables form the first slide.
	for (auto shape : shape_tems)
	{
		presentation->GetSlides()->GetItem(0)->GetShapes()->Remove(shape);
	}

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
