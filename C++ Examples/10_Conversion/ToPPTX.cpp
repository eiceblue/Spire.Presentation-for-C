#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ToPPTX.ppt";
	std::wstring outputFile = OutputPath"ToPPTX.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save the document to HTML format
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

	delete ppt;

}
