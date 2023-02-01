#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"EmbedExcelAsOLE.xlsx";
	std::wstring outputFile = OutputPath"EmbedExcelAsOLE.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();

	std::wstring imageFile = DataPath"EmbedExcelAsOLE.png";
	IImageData* oleImage = ppt->GetImages()->Append(new Stream(imageFile.c_str()));

	RectangleF* rec = new RectangleF(80, 60, oleImage->GetWidth(), oleImage->GetHeight());
	//Insert an OLE object to presentation based on the Excel data
	Spire::Presentation::IOleObject* oleObject = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendOleObject(L"excel", new Stream(inputFile.c_str()), rec);
	oleObject->GetSubstituteImagePictureFillFormat()->GetPicture()->SetEmbedImage(oleImage);
	oleObject->SetProgId(L"Excel.Sheet.12");

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;

}
