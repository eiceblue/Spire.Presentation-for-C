#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_1.pptx";
	std::wstring outputFile = OutputPath"AddImageWatermark.pptx";

	//Create a PowerPoint document.
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the image you want to add as image watermark.
	Stream* fileStream = new Stream(DataPath"Logo.png");

	IImageData* image = presentation->GetImages()->Append(fileStream);
	fileStream->Close();

	//Set the properties of SlideBackground, and then fill the image as watermark.
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->SetType(BackgroundType::Custom);
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Picture);
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->GetFill()->GetPictureFill()->SetFillType(PictureFillType::Stretch);
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->GetFill()->GetPictureFill()->GetPicture()->SetEmbedImage(image);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;
}
