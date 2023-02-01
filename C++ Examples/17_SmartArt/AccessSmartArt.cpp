#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"SmartArt.pptx";
	std::wstring outputFile = OutputPath"AccessSmartArt.txt";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	std::wstring* content = new std::wstring();

	content->append(L"Access SmartArt nodes.");
	content->append(L"\r\nHere is the SmartArt node parameters details:");
	ISmartArtNode* node;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		IShape* shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (dynamic_cast<ISmartArt*>(shape) != nullptr)
		{
			//Get the SmartArt
			ISmartArt* sa = dynamic_cast<ISmartArt*>(shape);

			ISmartArtNodeCollection* nodes = sa->GetNodes();

			//Traverse through all nodes inside SmartArt
			for (int i = 0; i < nodes->GetCount(); i++)
			{
				//Access SmartArt node at index i
				node = nodes->GetItem(i);
				std::wstring text = node->GetTextFrame()->GetText();
				//Print the SmartArt node parameters
				content->append(L"\r\nNode text = " + text + L", Node level = " + std::to_wstring(node->GetLevel()) + L", Node Position = " + std::to_wstring(node->GetPosition()));
			}
		}
	}
	//Save the file
	std::wofstream write(outputFile);
	write << content->c_str();
	write.close();
	delete presentation;
}
