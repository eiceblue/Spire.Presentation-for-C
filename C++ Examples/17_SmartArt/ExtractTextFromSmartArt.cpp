#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ExtractTextFromSmartArt.pptx";
	std::wstring outputFile = OutputPath"ExtractTextFromSmartArt.txt";

	//Create PPT document
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Traverse through all the slides of the PPT file and find the SmartArt shapes.
	wofstream outFile(outputFile, ios::out);
	outFile << "Below is extracted text from SmartArt:" << endl;

	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		for (int j = 0; j < presentation->GetSlides()->GetItem(i)->GetShapes()->GetCount(); j++)
		{
			if (dynamic_cast<ISmartArt*>(presentation->GetSlides()->GetItem(i)->GetShapes()->GetItem(j)) != nullptr)
			{
				ISmartArt* smartArt = dynamic_cast<ISmartArt*>(presentation->GetSlides()->GetItem(i)->GetShapes()->GetItem(j));

				//Extract text from SmartArt and append to the StringBuilder object.
				for (int k = 0; k < smartArt->GetNodes()->GetCount(); k++)
				{
					outFile << smartArt->GetNodes()->GetItem(k)->GetTextFrame()->GetText() << endl;
				}
			}
		}
	}
	//Save to file.
	outFile.close();
	delete presentation;
}
