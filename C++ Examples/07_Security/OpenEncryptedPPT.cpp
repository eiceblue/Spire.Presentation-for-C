#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"OpenEncryptedPPT.pptx";
	std::wstring outputFile = OutputPath"OpenEncryptedPPT.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the PPT with password
	presentation->LoadFromFile(inputFile.c_str(), FileFormat::Pptx2010, L"123456");

	//Save as a new PPT with original password
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;

}
