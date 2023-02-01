#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"sample.dps";
	std::wstring outputFile = OutputPath"LoadSaveDPSAndDPT.dps";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str(), FileFormat::Dps);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Dps);
	delete ppt;
}
