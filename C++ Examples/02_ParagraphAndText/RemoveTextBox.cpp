#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"TextBoxTemplate.pptx";
	std::wstring outputFile = OutputPath"RemoveTextBox.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = ppt->GetSlides()->GetItem(0);
	//Traverse all the shapes in slide
	for (int i = 0; i < slide->GetShapes()->GetCount();)
	{
		//Remove all shapes
		IAutoShape* shape = dynamic_cast<IAutoShape*>(slide->GetShapes()->GetItem(i));
		slide->GetShapes()->Remove(shape);
	}

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
