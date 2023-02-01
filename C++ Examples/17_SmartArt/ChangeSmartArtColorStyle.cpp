#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AddSmartArtNode.pptx";
	std::wstring outputFile = OutputPath"ChangeSmartArtColorStyle.pptx";

	//Create PPT document
	Presentation* presentation = new Presentation();
	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);

		if (dynamic_cast<ISmartArt*>(shape) != nullptr)
		{
			//Get the SmartArt and collect nodes
			ISmartArt* smartArt = dynamic_cast<ISmartArt*>(shape);
			// Check SmartArt color type
			if (smartArt->GetColorStyle() == SmartArtColorType::ColoredFillAccent1)
			{
				// Change SmartArt color type
				smartArt->SetColorStyle(SmartArtColorType::ColorfulAccentColors);
			}
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;

}
