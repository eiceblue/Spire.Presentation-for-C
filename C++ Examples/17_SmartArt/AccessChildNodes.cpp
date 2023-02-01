#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SmartArt.pptx";
	std::wstring outputFile = OutputPath"AccessChildNode.txt";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	std::wstring* content = new std::wstring();

	content->append(L"Access SmartArt child nodes.");
	content->append(L"\r\nHere is the SmartArt child node parameters details:");
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ISmartArt*>(shape) != nullptr)
		{
			//Get the SmartArt and collect nodes
			ISmartArt* sa = dynamic_cast<ISmartArt*>(shape);
			ISmartArtNodeCollection* nodes = sa->GetNodes();

			int position = 0;
			//Access the parent node at position 0
			ISmartArtNode* node = nodes->GetItem(position);
			ISmartArtNode* childnode;
			//Traverse through all child nodes inside SmartArt
			for (int i = 0; i < node->GetChildNodes()->GetCount(); i++)
			{
				//Access SmartArt child node at index i
				childnode = node->GetChildNodes()->GetItem(i);
				std::wstring text = childnode->GetTextFrame()->GetText();
				//Print the SmartArt child node parameters       
				content->append(L"\r\nNode text = " + text + L", Node level = " + std::to_wstring(childnode->GetLevel()) + L", Node Position = " + std::to_wstring(childnode->GetPosition()));
			}
		}
	}
	//Save the file
	std::wofstream write(outputFile);
	write << content->c_str();
	write.close();
	delete presentation;
}
