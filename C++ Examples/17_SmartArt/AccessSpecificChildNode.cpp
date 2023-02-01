#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SmartArt.pptx";
	std::wstring outputFile = OutputPath"AccessSpecificChildNode.txt";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);
	outFile << "Access SmartArt child node at specific position." << endl;
	outFile << "Here is the SmartArt child node parameters details:" << endl;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ISmartArt*>(shape) != nullptr)
		{
			//Get the SmartArt
			ISmartArt* sa = dynamic_cast<ISmartArt*>(shape);

			//Get SmartArt node collection 
			ISmartArtNodeCollection* nodes = sa->GetNodes();

			//Access SmartArt node at index 0
			ISmartArtNode* node = nodes->GetItem(0);

			//Access SmartArt child node at index 1
			ISmartArtNode* childNode = node->GetChildNodes()->GetItem(1);

			//Print the SmartArt child node parameters
			outFile << "Node text = " << childNode->GetTextFrame()->GetText() << ", Node level = " << childNode->GetLevel() << ", Node Position = " << childNode->GetPosition() << endl;
		}
	}
	//Save the file
	outFile.close();
	delete presentation;
}
