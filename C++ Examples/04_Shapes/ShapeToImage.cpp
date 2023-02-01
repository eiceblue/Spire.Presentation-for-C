#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ShapeToImage.pptx";
	std::wstring outputFile = OutputPath"Image/ShapeToImage/";

	//Create a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());
	ISlide* slide = presentation->GetSlides()->GetItem(0);
	ShapeCollection* shapes = slide->GetShapes();
	for (int i = 0; i < shapes->GetCount(); i++)
	{
		Stream* image = shapes->SaveAsImage(i);
		std::wstring filename = outputFile + L"//ShapeToImage-" + to_wstring(i) + L".png";
		image->Save(filename.c_str());
		image->Dispose();
	}
	delete presentation;
}
