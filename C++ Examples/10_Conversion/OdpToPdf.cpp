#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"toPdf.odp";
	std::wstring outputFile = OutputPath"OdpToPdf.pdf";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	delete ppt;
}
