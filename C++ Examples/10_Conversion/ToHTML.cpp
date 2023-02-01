#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Conversion.pptx";
	std::wstring outputFile = OutputPath"ToHTML.html";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save the document to HTML format
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Html);

	delete ppt;
}
