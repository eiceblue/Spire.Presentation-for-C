#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Template_Ppt_4.pptx";
	std::wstring outputFile = OutputPath"RemoveEncryption.pptx";

	//Create a PowerPoint document.
	Presentation* presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str(), L"123456");

	//Remove encryption.
	presentation->RemoveEncryption();

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete presentation;

}
