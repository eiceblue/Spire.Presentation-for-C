#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ToPDF.pptx";
	std::wstring outputFile = OutputPath"ToPDF.pdf";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save the document to HTML format
	ppt->SaveToFile(outputFile.c_str(), FileFormat::PDF);

	delete ppt;
}
