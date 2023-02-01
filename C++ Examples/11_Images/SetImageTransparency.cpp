#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"SetImageTransparency.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	//Load an image
	std::wstring imageFile = DataPath"Logo.png";
	std::ifstream inputf(imageFile.c_str(), std::ios::in | std::ios::binary);
	Stream* stream = new Stream(inputf);
	IImageData* imageData = ppt->GetImages()->Append(stream);

	//Add the image in document
	RectangleF* rect = new RectangleF(200, 100, imageData->GetWidth(), imageData->GetHeight());
	//Add a shape
	IAutoShape* shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rect);
	shape->GetLine()->SetFillType(FillFormatType::None);
	//Fill shape with image
	shape->GetFill()->SetFillType(FillFormatType::Picture);
	shape->GetFill()->GetPictureFill()->GetPicture()->SetUrl(imageFile.c_str());
	shape->GetFill()->GetPictureFill()->SetFillType(PictureFillType::Stretch);
	//Set transparency on image
	shape->GetFill()->GetPictureFill()->GetPicture()->SetTransparency(50);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
