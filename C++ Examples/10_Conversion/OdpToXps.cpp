#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"toPdf.odp";
	std::wstring outputFile = OutputPath"OdpToXps.xps";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::XPS);
	delete ppt;
}
