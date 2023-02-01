#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"InsertImage.pptx";
	std::wstring outputFile = OutputPath"InsertImage.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//EMF file path
	std::wstring ImageFile = DataPath"InsertImage.png";

	//Define image size
	RectangleF* rect = new RectangleF(ppt->GetSlideSize()->GetSize()->GetWidth() / 2 - 280, 140, 120, 120);

	//Append the EMF in slide
	IEmbedImage* image = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	image->GetLine()->SetFillType(FillFormatType::None);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;

}
