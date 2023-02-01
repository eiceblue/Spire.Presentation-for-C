#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AddSmartArtNode.pptx";
	std::wstring outputFile = OutputPath"ChangeSmartArtShapeStyle.pptx";

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
			//Check SmartArt style
			if (smartArt->GetStyle() == SmartArtStyleType::SimpleFill)
			{
				//Change SmartArt Style
				smartArt->SetStyle(SmartArtStyleType::Cartoon);
			}
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
