#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ExtractImage.pptx";
	std::wstring outputFile = OutputPath"Image/ExtractImage/";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	for (int i = 0; i < ppt->GetImages()->GetCount(); i++)
	{
		std::wstring ImageName = outputFile + L"Images_" + to_wstring(i) + L".png";
		Stream* image = ppt->GetImages()->GetItem(i)->GetImage();
		image->Save(ImageName.c_str());
		delete image;
	}
	delete ppt;

}
