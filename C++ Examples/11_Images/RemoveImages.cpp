#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"RemoveImages.pptx";
	std::wstring outputFile = OutputPath"RemoveImages.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	ISlide* slide = ppt->GetSlides()->GetItem(0);

	for (int i = slide->GetShapes()->GetCount() - 1; i >= 0; i--)
	{
		//It is the SlidePicture object
		if (dynamic_cast<IEmbedImage*>(slide->GetShapes()->GetItem(i)) != nullptr)
		{
			slide->GetShapes()->RemoveAt(i);
		}
	}
	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
