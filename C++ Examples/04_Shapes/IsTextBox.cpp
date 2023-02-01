#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"IsTextboxSample.pptx";
	std::wstring outputFile = OutputPath"IsTextBox.txt";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);

	for (int l = 0; l < ppt->GetSlides()->GetCount(); l++)
	{
		ISlide* slide = ppt->GetSlides()->GetItem(l);
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			IShape* shape = slide->GetShapes()->GetItem(s);
			if (dynamic_cast<IAutoShape*>(shape) != nullptr)
			{
				//Judge if the shape is textbox
				bool isTextbox = shape->GetIsTextBox();
				outFile << (isTextbox ? "shape is text box" : "shape is not text box") << endl;
			}
		}
	}
	outFile.close();
	delete ppt;
}
