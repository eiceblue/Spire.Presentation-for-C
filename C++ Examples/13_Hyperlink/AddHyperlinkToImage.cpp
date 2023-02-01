#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_5.pptx";
	std::wstring outputFile = OutputPath"AddHyperlinkToImage.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide.
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Add image to slide.
	RectangleF* rect = new RectangleF(480, 350, 160, 160);
	std::wstring inputFile1 = DataPath"Logo1.png";
	IEmbedImage* image = slide->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, inputFile1.c_str(), rect);

	//Add hyperlink to the image.
	ClickHyperlink* hyperlink = new ClickHyperlink(L"https://www.e-iceblue.com");
	image->SetClick(hyperlink);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;

}
