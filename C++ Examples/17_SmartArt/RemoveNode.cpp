#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"RemoveNode.pptx";
	std::wstring outputFile = OutputPath"RemoveNode.pptx";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the SmartArt and collect nodes
	ISmartArt* sa = dynamic_cast<ISmartArt*>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	ISmartArtNodeCollection* nodes = sa->GetNodes();

	//Remove the node to specific position
	nodes->RemoveNodeByPosition(2);

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
