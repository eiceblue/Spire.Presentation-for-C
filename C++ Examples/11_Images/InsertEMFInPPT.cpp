#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"BlankSample_N.pptx";
	std::wstring outputFile = OutputPath"InsertEMFInPPT.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//EMF file path
	std::wstring ImageFile = DataPath"InsertEMF.emf";

	//Define image size
	RectangleF* rect = new RectangleF(100, 100, 719 / 1.5, 539 / 1.5);
	//Append the EMF in slide
	IEmbedImage* image = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	image->GetLine()->SetFillType(FillFormatType::None);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
