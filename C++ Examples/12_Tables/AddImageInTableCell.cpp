#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"AddImageInTableCell.pptx";
	std::wstring outputFile = OutputPath"AddImageInTableCell.pptx";

	//Load a PPT document
	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first shape
	ITable* table = dynamic_cast<ITable*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Load the image and insert it into table cell

	Stream* stream = new Stream(DataPath"PresentationIcon.png");
	IImageData* pptImg = ppt->GetImages()->Append(stream);
	stream->Close();

	table->GetItem(1, 1)->GetFillFormat()->SetFillType(FillFormatType::Picture);
	table->GetItem(1, 1)->GetFillFormat()->GetPictureFill()->GetPicture()->SetEmbedImage(pptImg);
	table->GetItem(1, 1)->GetFillFormat()->GetPictureFill()->SetFillType(PictureFillType::Stretch);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
