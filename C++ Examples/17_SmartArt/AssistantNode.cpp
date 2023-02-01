#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AddSmartArtNode.pptx";
	std::wstring outputFile = OutputPath"AssistantNode.pptx";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());
	ISmartArtNode* node;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ISmartArt*>(shape) != nullptr)
		{
			//Get the SmartArt and collect nodes
			ISmartArt* smartArt = dynamic_cast<ISmartArt*>(shape);

			ISmartArtNodeCollection* nodes = smartArt->GetNodes();

			//Traverse through all nodes inside SmartArt
			for (int i = 0; i < nodes->GetCount(); i++)
			{
				//Access SmartArt node at index i
				node = nodes->GetItem(i);
				// Check if node is assitant node
				if (!node->GetIsAssistant())
				{
					//Set node as assitant node
					node->SetIsAssistant(true);
				}
			}
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
