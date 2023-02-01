#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"EmbedZipIntoPPT.pptx";
	std::wstring inputFile_z = DataPath"test.zip";
	std::wstring inputFile_I = DataPath"icon.png";
	std::wstring outputFile = OutputPath"EmbedZipIntoPPT.pptx";

	//Create a Presentaion document
	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Load a zip object
	Stream* zipStream = new Stream(inputFile_z.c_str());

	RectangleF* rec = new RectangleF(80, 60, 100, 100);

	//Insert the zip object to presentation
	Spire::Presentation::IOleObject* ole = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendOleObject(L"zip", zipStream, rec);
	ole->SetProgId(L"Package");
	Stream* image = new Stream(inputFile_I.c_str());
	IImageData* oleImage = ppt->GetImages()->Append(image);
	ole->GetSubstituteImagePictureFillFormat()->GetPicture()->SetEmbedImage(oleImage);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;

}
