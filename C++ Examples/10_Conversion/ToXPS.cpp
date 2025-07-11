#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Conversion.pptx";
	wstring outputFile = OUTPUTPATH"ToXPS.xps";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save the document to HTML format
	ppt->SaveToFile(outputFile.c_str(), FileFormat::XPS);

}
