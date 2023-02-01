#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring outputFile = OutputPath"CreateSlideMasterAndApply.pptx";

	//Create an instance of presentation document
	Presentation* ppt = new Presentation();

	ppt->GetSlideSize()->SetType(SlideSizeType::Screen16x9);

	//Add slides
	for (int i = 0; i < 4; i++)
	{
		ppt->GetSlides()->Append();
	}

	//Get the first default slide master
	IMasterSlide* first_master = ppt->GetMasters()->GetItem(0);

	//Append another slide master
	ppt->GetMasters()->AppendSlide(first_master);
	IMasterSlide* second_master = ppt->GetMasters()->GetItem(1);

	//Set different background image for the two slide masters
	std::wstring pic1 = DataPath"bg.png";
	std::wstring pic2 = DataPath"Setbackground.png";
	//The first slide master
	RectangleF* rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	first_master->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Picture);
	IEmbedImage* image1 = first_master->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, pic1.c_str(), rect);
	first_master->GetSlideBackground()->GetFill()->GetPictureFill()->GetPicture()->SetEmbedImage(image1->GetPictureFill()->GetPicture()->GetEmbedImage());
	//The second slide master
	second_master->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Picture);
	IEmbedImage* image2 = second_master->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, pic2.c_str(), rect);
	second_master->GetSlideBackground()->GetFill()->GetPictureFill()->GetPicture()->SetEmbedImage(image2->GetPictureFill()->GetPicture()->GetEmbedImage());

	//Apply the first master with layout to the first slide
	ppt->GetSlides()->GetItem(0)->SetLayout(first_master->GetLayouts()->GetItem(1));

	//Apply the second master with layout to other slides
	for (int i = 1; i < ppt->GetSlides()->GetCount(); i++)
	{
		ppt->GetSlides()->GetItem(i)->SetLayout(second_master->GetLayouts()->GetItem(8));
	}

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
