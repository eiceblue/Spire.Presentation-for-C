#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AddImageInMaster.pptx";
	std::wstring outputFile = OutputPath"AddImageInMaster.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the master collection
	IMasterSlide* master = presentation->GetMasters()->GetItem(0);

	//Append image to slide master
	std::wstring image = DataPath"Logo.png";
	RectangleF* rff = new RectangleF(40, 40, 90, 90);
	IEmbedImage* pic = master->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, image.c_str(), rff);
	pic->GetLine()->GetFillFormat()->SetFillType(FillFormatType::None);

	//Add new slide to presentation
	presentation->GetSlides()->Append();

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;
}
