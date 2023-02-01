#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"FillShapeWithPicture.pptx";
	std::wstring outputFile = OutputPath"FillShapeWithPicture.pptx";

	//Load a PPT document
	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first shape and set the style to be Gradient
	IAutoShape* shape = dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Fill the shape with picture
	std::wstring picUrl = DataPath"backgroundImg.png";
	shape->GetFill()->SetFillType(FillFormatType::Picture);
	shape->GetFill()->GetPictureFill()->GetPicture()->SetUrl(picUrl.c_str());
	shape->GetFill()->GetPictureFill()->SetFillType(PictureFillType::Stretch);

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
